"""
Build the public nursery website from a read-only copy of plant-index.xlsx.

    python build_site.py                  (uses site.json)
    python build_site.py other-copy.xlsx  (use a different workbook)
    python build_site.py --force          (re-make every photo, even unchanged ones)

Output: a folder of plain HTML/CSS/JS files (default: site/). Open site/index.html
to preview; upload the whole folder to any web host. The workbook is only READ.

Windows note: this script never deletes the whole output folder. Deleting a folder
that Explorer, a browser, OneDrive/Dropbox or antivirus is looking at is what
causes "[WinError 5] Access is denied". Files are overwritten in place instead.
"""
import argparse
import json
import os
import re
import shutil
import stat
import time
import unicodedata
from datetime import date, datetime
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup
from openpyxl import load_workbook
from PIL import Image, ImageOps

BASE = Path(__file__).resolve().parent
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
FULL_SIZE, THUMB_SIZE, HERO_SIZE = 1200, 480, 1920   # longest side in pixels
SIZE_WORDS = {"QT": "Quart", "6PK": "6-pack"}
STATUS_CLASS = {"In stock": "ok", "Low stock": "low", "Check stock": "check", "": "check"}


# ---------------------------------------------------------------- safe file output
def _retry(fn, tries=6):
    """Run fn(); if Windows says 'access denied' (file briefly locked), wait and retry."""
    for i in range(tries):
        try:
            return fn()
        except PermissionError:
            if i == tries - 1:
                raise
            time.sleep(0.4 * (i + 1))


def _make_writable(path):
    if path.exists():
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)


def write_file(dest, data):
    """Write bytes/str to dest via a temp file, so a locked or read-only dest can't break the build."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.name + ".tmp")
    tmp.write_bytes(data.encode("utf-8") if isinstance(data, str) else data)
    _make_writable(dest)
    _retry(lambda: os.replace(tmp, dest))


def remove_file(path):
    try:
        _make_writable(path)
        _retry(path.unlink)
    except OSError as e:
        print(f"  ! could not remove {path.name} ({e}); leaving it")


def prune(folder, keep, pattern="*"):
    """Delete files in folder (recursively) that aren't in `keep`, then any empty sub-folders.
    Never raises: anything that is locked is simply left for next time."""
    if not folder.exists():
        return
    for f in folder.rglob(pattern):
        if f.is_file() and f.resolve() not in keep:
            remove_file(f)
    for d in sorted((d for d in folder.rglob("*") if d.is_dir()), reverse=True):
        try:
            d.rmdir()                     # only succeeds if empty
        except OSError:
            pass


# ---------------------------------------------------------------- text helpers
def text(v):
    return "" if v is None else str(v).strip()


def slugify(s):
    s = unicodedata.normalize("NFKD", text(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def nicecase(v):
    s = text(v)
    if s and s == s.upper():
        s = re.sub(r"[A-Za-z]+(?:['\u2019][A-Za-z]+)*", lambda m: m.group(0).capitalize(), s)
    return s


def botanical_html(v):
    """Italicise the species part; leave the 'Cultivar' in quotes upright."""
    s = text(v)
    m = re.search(r"[\u2018\u2019'\u201c\"]", s)
    if not m:
        return Markup("<i>%s</i>") % s
    return Markup("<i>%s</i> %s") % (s[:m.start()].rstrip(), s[m.start():])


def size_label(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f"{v:g} gal"
    s = text(v)
    return SIZE_WORDS.get(s.upper(), s)


def money(v):
    try:
        return f"${float(v):,.2f}"
    except (TypeError, ValueError):
        return ""


def botanical_sort_key(botanical, common=""):
    """Alphabetical by botanical name; ignores case, quote marks and stray spaces so
    Acer palmatum 'Bloodgood' and Acer palmatum \u2018Bloodgood\u2019 sort the same way."""
    b = re.sub(r"[\u2018\u2019\u201c\u201d'\"]", "", text(botanical))
    return (re.sub(r"\s+", " ", b).strip().casefold(), text(common).casefold())


def size_key(r):
    v = r.get("SIZE")
    return (0, float(v), "") if isinstance(v, (int, float)) else (1, 0, text(v).lower())


def sun_flags(exposure):
    """F = full sun, P = part sun/part shade, S = shade."""
    e = text(exposure).upper()
    flags = ""
    if "FULL" in e or e.startswith("SUN"):
        flags += "F"
    if "PART" in e:
        flags += "P"
    if re.search(r"(^|-)SHADE", e):
        flags += "S"
    return flags


def parse_images(v):
    out = []
    for part in re.split(r"[;\n\r]+", text(v)):
        p = re.sub(r"^(\./|/)+", "", part.strip().replace("\\", "/"))
        if p.lower().startswith("photos/"):
            p = p[7:]
        if p and p not in out:
            out.append(p)
    return out


# ---------------------------------------------------------------- photos
def make_resized(src, dest, size):
    """Save a shrunken copy of src to dest (via write_file), keeping src's modified-time."""
    ext = src.suffix.lower()
    fmt = {".jpg": "JPEG", ".jpeg": "JPEG", ".png": "PNG", ".webp": "WEBP", ".gif": "GIF"}[ext]
    with Image.open(src) as im:
        im = ImageOps.exif_transpose(im)
        if fmt == "JPEG" and im.mode != "RGB":
            im = im.convert("RGB")
        im.thumbnail((size, size))
        tmp = dest.with_name(dest.name + ".img")
        dest.parent.mkdir(parents=True, exist_ok=True)
        im.save(tmp, format=fmt, **({"quality": 82, "optimize": True} if fmt in ("JPEG", "WEBP") else {}))
    write_file(dest, tmp.read_bytes())
    tmp.unlink()
    mtime = int(src.stat().st_mtime)
    os.utime(dest, (mtime, mtime))       # lets the next build skip this photo if the source is unchanged


def publish_photo(rel, photo_dir, out_photos, cache, force):
    """Returns {'full': url, 'thumb': url} for one photo, or None if it's missing or can't be processed."""
    if rel in cache:
        return cache[rel]
    src = (photo_dir / rel).resolve()
    cache[rel] = None
    if photo_dir.resolve() not in src.parents or not src.is_file() or src.suffix.lower() not in IMAGE_EXTS:
        print(f"  ! photo not found, skipped: {rel}")
        return None
    full, thumb = out_photos / rel, out_photos / "thumbs" / rel
    try:
        for dest, size in ((full, FULL_SIZE), (thumb, THUMB_SIZE)):
            if force or not dest.is_file() or int(dest.stat().st_mtime) != int(src.stat().st_mtime):
                make_resized(src, dest, size)
    except Exception as e:                # one bad photo must not stop the whole build
        print(f"  ! could not process {rel}: {e}")
        return None
    cache[rel] = {"full": f"photos/{rel}", "thumb": f"photos/thumbs/{rel}"}
    return cache[rel]


def build_hero(cfg, photo_dir, out_dir, force):
    """Resize the homepage hero image to static/hero.jpg. Returns its site-relative URL, or None
    if no hero_image is set or the file is missing (the hero then shows a plain grey-green panel)."""
    rel = text(cfg.get("hero_image"))
    if not rel:
        return None
    src = (photo_dir / parse_images(rel)[0]).resolve() if parse_images(rel) else None
    if not src or photo_dir.resolve() not in src.parents or not src.is_file() or src.suffix.lower() not in IMAGE_EXTS:
        print(f"  ! hero image not found, using plain overlay: {rel}")
        return None
    dest = out_dir / "static" / "hero.jpg"
    try:
        if force or not dest.is_file() or int(dest.stat().st_mtime) != int(src.stat().st_mtime):
            with Image.open(src) as im:
                im = ImageOps.exif_transpose(im).convert("RGB")
                im.thumbnail((HERO_SIZE, HERO_SIZE))
                tmp = dest.with_name(dest.name + ".img")
                dest.parent.mkdir(parents=True, exist_ok=True)
                im.save(tmp, format="JPEG", quality=80, optimize=True)
            write_file(dest, tmp.read_bytes())
            tmp.unlink()
            os.utime(dest, (int(src.stat().st_mtime),) * 2)
    except Exception as e:
        print(f"  ! could not process hero image: {e}")
        return None
    return "static/hero.jpg"


# ---------------------------------------------------------------- data
def load_rows(xlsx):
    wb = load_workbook(xlsx, read_only=True, data_only=True)   # read-only: never modifies the file
    it = wb.worksheets[0].iter_rows(values_only=True)
    headers = [text(h) for h in next(it)]
    rows = [dict(zip(headers, vals)) for vals in it if any(v not in (None, "") for v in vals)]
    wb.close()
    return rows


def date_value(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return text(v)


def build_plants(rows, cfg, photo_dir, out_dir, force):
    labels = dict(cfg["status_labels"])
    if cfg.get("blank_status_label"):          # optional: list plants whose STOCK NOTE is blank
        labels[""] = cfg["blank_status_label"]
    status_order = list(labels)                # best first

    groups = {}
    for r in rows:
        if text(r.get("STOCK NOTE")) not in labels:
            continue                           # hidden: blank, Not available, anything unlisted
        if not text(r.get("COMMON NAME")) and not text(r.get("BOTANICAL NAME")):
            continue
        slug = f"{slugify(r.get('BOTANICAL NAME'))}--{slugify(r.get('COMMON NAME'))}"
        groups.setdefault(slug, []).append(r)

    cache, plants = {}, []
    for slug, grp in groups.items():
        grp.sort(key=size_key)
        first = lambda field: next((r[field] for r in grp if text(r.get(field))), None)

        photos = [p for p in (publish_photo(rel, photo_dir, out_dir / "photos", cache, force)
                              for rel in parse_images(first("IMAGE"))) if p]
        sizes = [{"size": size_label(r.get("SIZE")) or "-", "price": money(r.get("PRICE")),
                  "status": labels[text(r.get("STOCK NOTE"))],
                  "cls": STATUS_CLASS.get(text(r.get("STOCK NOTE")), "check")} for r in grp]
        prices = [float(r["PRICE"]) for r in grp if isinstance(r.get("PRICE"), (int, float))]
        best = min((text(r.get("STOCK NOTE")) for r in grp), key=status_order.index)
        category = text(first("SUPERTYPE")) or cfg["category_other"]
        subtype = text(first("TYPE"))

        plants.append({
            "slug": slug,
            "common": nicecase(first("COMMON NAME")) or text(first("BOTANICAL NAME")),
            "botanical": text(first("BOTANICAL NAME")),
            "botanical_html": botanical_html(first("BOTANICAL NAME")),
            "category": category, "type": subtype, "subtype": subtype,
            "last_updated": date_value(first("LAST CHECK")),
            "sun": nicecase(first("EXPOSURE")), "sun_flags": sun_flags(first("EXPOSURE")),
            "hxw": text(first("HxW")),
            "traits": [t for t in (text(first("CN / DT")), text(first("FOLIAGE TYPE"))) if t],
            "photos": photos, "sizes": sizes,
            "from_price": money(min(prices)) if prices else "",
            "status": labels[best], "cls": STATUS_CLASS.get(best, "check"),
        })
    plants.sort(key=lambda p: botanical_sort_key(p["botanical"], p["common"]))
    return plants


# ---------------------------------------------------------------- output
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?", help="workbook to use instead of the one in site.json")
    ap.add_argument("--force", action="store_true", help="re-make every photo")
    args = ap.parse_args()

    cfg = json.loads((BASE / "site.json").read_text(encoding="utf-8"))
    xlsx = Path(args.xlsx).resolve() if args.xlsx else BASE / cfg["source_xlsx"]
    photo_dir = BASE / cfg["photo_dir"]
    out_dir = BASE / cfg["output_dir"]

    print(f"Reading {xlsx}")
    rows = load_rows(xlsx)
    out_dir.mkdir(exist_ok=True)

    plants = build_plants(rows, cfg, photo_dir, out_dir, args.force)
    hero_url = build_hero(cfg, photo_dir, out_dir, args.force)

    today = date.today()
    cfg["updated"] = f"{today:%B} {today.day}, {today.year}"
    cfg["maps_url"] = "https://www.google.com/maps/search/?api=1&query=" + \
        re.sub(r"\s+", "+", ", ".join(cfg.get("address", [])))

    counts = {}
    for p in plants:
        counts[p["category"]] = counts.get(p["category"], 0) + 1
    order = list(cfg.get("supertype_order", []))
    categories = [(c, counts[c]) for c in order if c in counts]
    categories += sorted([(c, n) for c, n in counts.items() if c not in order], key=lambda x: x[0].casefold())

    featured = [p for p in plants if p["photos"]]
    featured.sort(key=lambda p: p["cls"] != "ok")          # in-stock first (stable sort)
    featured = featured[:cfg["featured_count"]]

    env = Environment(loader=FileSystemLoader(BASE / "site_templates"), autoescape=select_autoescape(["html"]))
    written = set()

    def out(rel, data):
        write_file(out_dir / rel, data)
        written.add((out_dir / rel).resolve())

    def render(name, dest, root, **ctx):
        out(dest, env.get_template(name).render(site=cfg, root=root, **ctx))

    # static files (copied as plain bytes, so no read-only flags come along)
    for f in (BASE / "site_static").rglob("*"):
        if f.is_file():
            out(Path("static") / f.relative_to(BASE / "site_static"), f.read_bytes())

    if hero_url:
        written.add((out_dir / hero_url).resolve())     # keep it safe from the static/ tidy-up below
    render("home.html", "index.html", "", nav="home", featured=featured, categories=categories,
           plant_count=len(plants), hero_url=hero_url)
    render("catalog.html", "plants.html", "", nav="plants", categories=categories)
    render("about.html", "about.html", "", nav="about")
    render("visit.html", "visit.html", "", nav="visit")
    for p in plants:
        render("plant.html", f"plant/{p['slug']}.html", "../", nav="plants", p=p)

    # Compact search data, loaded by plants.html as a plain script (works from file:// too).
    data = [{"s": p["slug"], "n": p["common"], "b": p["botanical"], "c": p["category"], "sub": p["subtype"], "d": p["last_updated"],
             "u": p["sun_flags"], "t": p["traits"], "p": p["from_price"], "st": p["status"],
             "k": p["cls"], "i": p["photos"][0]["thumb"] if p["photos"] else "",
             "h": " ".join([p["common"], p["botanical"], p["category"], p["subtype"], p["type"], p["sun"],
                            *p["traits"]]).lower()} for p in plants]
    out("static/plants-data.js",
        "window.PLANTS=" + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";")

    # Tidy up: remove pages and photos for plants that are no longer listed (skips anything locked).
    for p in plants:
        for ph in p["photos"]:
            written.update(((out_dir / ph["full"]).resolve(), (out_dir / ph["thumb"]).resolve()))
    prune(out_dir / "plant", written, "*.html")
    prune(out_dir / "static", written)

    with_photo = sum(1 for p in plants if p["photos"])
    print(f"Done: {len(plants)} plants listed ({with_photo} with photos) from {len(rows)} workbook rows.")
    print(f"Open {out_dir / 'index.html'} to preview.")


if __name__ == "__main__":
    main()
