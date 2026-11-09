from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, abort, jsonify
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, date
from pathlib import Path
from werkzeug.utils import secure_filename
import threading
import math
import re
import unicodedata
import subprocess
import sys

BASE = Path(__file__).resolve().parent
XLSX = BASE / "plant-index.xlsx"
SHEET = None
LOCK = threading.Lock()
PAGE_SIZE = 50

# --- Plant landing pages & photos -------------------------------------------
# Photos live in this folder. The IMAGE column in the workbook holds the path
# *relative to this folder*, e.g.  roses/knockout-red.jpg
PHOTO_DIR = BASE / "photos"
IMAGE_FIELD = "IMAGE"            # column heading in the workbook
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
# Internal columns shown only in the "Staff details" panel of a landing page.
STAFF_FIELDS = ["ROW", "#", "LAST CHECK", "TYPE", "SUPERTYPE"]
# Set to False before sharing landing pages with customers: hides the staff panel,
# the Edit buttons, and photo-path warnings.
SHOW_STAFF_DETAILS = True

# Quick "we have this plant" checkbox on the main results page. Ticking it writes
# today's date into this column; unticking undoes it. No need to open the record.
CHECK_FIELD = "LAST CHECK"

# Form layout: header name -> section. Headers not listed fall into "Other Fields",
# so the form keeps working if columns are added or renamed in the workbook.
FORM_SECTIONS = [
    ("Plant Identification", ["BOTANICAL NAME", "COMMON NAME", "SUPERTYPE", "TYPE"]),
    ("Inventory & Pricing", ["STOCK NOTE", "SIZE", "PRICE", "LAST CHECK", "ROW", "#"]),
    ("Growing Information", ["HxW", "EXPOSURE", "CN / DT", "FOLIAGE TYPE"]),
    ("Photo", [IMAGE_FIELD]),
]
WIDE_FIELDS = {"BOTANICAL NAME", "COMMON NAME"}
# Fields that suggest existing values from the workbook while still allowing free text.
SUGGEST_FIELDS = ["SUPERTYPE", "TYPE", "SIZE", "EXPOSURE", "CN / DT", "FOLIAGE TYPE"]

app = Flask(__name__)
app.secret_key = "local-plant-index"

def get_workbook():
    wb = load_workbook(XLSX)
    ws = wb[wb.sheetnames[0]]
    return wb, ws

def slugify(text):
    """URL-safe text. Curly and straight quotes compare equal, so
    'Bloodgood' and \u2018Bloodgood\u2019 end up on the same page."""
    text = unicodedata.normalize("NFKD", str(text or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")

def plant_slug(botanical, common):
    return f"{slugify(botanical)}--{slugify(common)}"

def clean_image_path(value):
    """Normalise the IMAGE cell to a path relative to PHOTO_DIR (or '')."""
    v = display_value(value).strip().replace("\\", "/")
    v = re.sub(r"^(\./|/)+", "", v)
    if v.lower().startswith("photos/"):
        v = v[7:]
    return v

def photo_files():
    """Image files currently in the photo library (for the form's type-ahead)."""
    if not PHOTO_DIR.is_dir():
        return []
    files = [f.relative_to(PHOTO_DIR).as_posix() for f in PHOTO_DIR.rglob("*")
             if f.is_file() and f.suffix.lower() in IMAGE_EXTS]
    return sorted(files, key=str.lower)[:5000]

def photo_exists(rel):
    if not rel:
        return False
    target = (PHOTO_DIR / rel).resolve()
    return PHOTO_DIR.resolve() in target.parents and target.is_file()

def headers_and_rows():
    wb, ws = get_workbook()
    headers = [str(ws.cell(1, c).value).strip() if ws.cell(1, c).value is not None else f"UNNAMED_{c}" for c in range(1, ws.max_column + 1)]
    rows = []
    for excel_row in range(2, ws.max_row + 1):
        values = [ws.cell(excel_row, c).value for c in range(1, ws.max_column + 1)]
        # Ignore completely blank rows.
        if not any(v is not None and v != "" for v in values):
            continue
        item = {"_row": excel_row}
        for i, h in enumerate(headers):
            item[h] = values[i]
        item["_slug"] = plant_slug(item.get("BOTANICAL NAME"), item.get("COMMON NAME"))
        rows.append(item)
    wb.close()
    return headers, rows

def display_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value)

def stock_options():
    """Allowed STOCK NOTE values, read from Sheet2 (the workbook's dropdown source)."""
    wb = load_workbook(XLSX)
    options = []
    if len(wb.sheetnames) > 1:
        ws2 = wb[wb.sheetnames[1]]
        for r in range(1, ws2.max_row + 1):
            v = ws2.cell(r, 1).value
            if v not in (None, ""):
                options.append(str(v))
    wb.close()
    valid = ["In stock", "Low stock", "Check stock", "Not available"]
    return valid

def parse_image_paths(value):
    """Read legacy single paths or semicolon/newline-separated image paths."""
    raw = display_value(value).replace("\r", "\n")
    parts = re.split(r"[;\n]+", raw)
    paths = []
    for part in parts:
        rel = clean_image_path(part)
        if rel and rel not in paths:
            paths.append(rel)
    return paths

def form_sections(headers):
    """Group (column_number, header) pairs into form sections."""
    index = {h: i for i, h in enumerate(headers, 1)}
    used, sections = set(), []
    for title, names in FORM_SECTIONS:
        fields = [(index[n], n) for n in names if n in index]
        used.update(n for n in names if n in index)
        if fields:
            sections.append((title, fields))
    other = [(i, h) for h, i in index.items() if h not in used and h is not None]
    if other:
        sections.append(("Other Fields", other))
    return sections

def suggestions(headers, rows):
    """Distinct existing values per field, for the form's type-ahead lists."""
    out = {}
    for h in SUGGEST_FIELDS:
        if h in headers:
            vals = {display_value(r[h]) for r in rows if r[h] not in (None, "")}
            out[h] = sorted(vals, key=lambda v: (not v.isdigit(), int(v) if v.isdigit() else 0, v.lower()))
    return out

def restore_stock_validation(ws, wb):
    """openpyxl drops Excel's newer-style dropdown on save; re-add it so the
    STOCK NOTE dropdown (source: Sheet2) keeps working in Excel."""
    if len(wb.sheetnames) < 2:
        return
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    if "STOCK NOTE" not in headers:
        return
    col = headers.index("STOCK NOTE") + 1
    from openpyxl.utils import get_column_letter
    letter = get_column_letter(col)
    last = wb[wb.sheetnames[1]].max_row
    target = f"'{wb.sheetnames[1]}'!$A$2:$A${last}"
    # Replace any earlier copy of this dropdown so repeated saves never stack duplicates.
    ws.data_validations.dataValidation = [
        d for d in ws.data_validations.dataValidation
        if not (d.type == "list" and (d.formula1 or "").replace("'", "") == target.replace("'", ""))
    ]
    dv = DataValidation(type="list", formula1=target,
                        allow_blank=True, showErrorMessage=False)
    dv.add(f"{letter}2:{letter}1048576")
    ws.add_data_validation(dv)

def save_form(excel_row=None):
    with LOCK:
        wb, ws = get_workbook()
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

        if excel_row is None:
            excel_row = ws.max_row + 1

        for c, header in enumerate(headers, 1):
            value = request.form.get(f"field_{c}", "")
            if header == "LAST CHECK" and value:
                try:
                    value = datetime.strptime(value, "%Y-%m-%d")
                except ValueError:
                    pass
            elif header == "PRICE" and value != "":
                try:
                    value = float(value)
                except ValueError:
                    pass
            elif header == "SIZE" and value != "":
                try:
                    value = int(value)
                except ValueError:
                    pass
            elif value == "":
                value = None
            ws.cell(excel_row, c).value = value

        restore_stock_validation(ws, wb)
        wb.save(XLSX)
        wb.close()

@app.template_filter("stock_class")
def stock_class(value):
    v = (value or "").strip().lower()
    if v in ("in stock", "in"):
        return "ok"
    if v == "low stock":
        return "low"
    if v == "not available":
        return "out"
    if v == "check stock":
        return "check"
    return "none"

@app.template_filter("nicecase")
def nicecase(value):
    """ALL CAPS common names -> Title Case (keeps apostrophes sane)."""
    s = display_value(value)
    if s and s == s.upper():
        s = re.sub(r"[A-Za-z]+(?:['\u2019][A-Za-z]+)*", lambda m: m.group(0).capitalize(), s)
    return s

@app.template_filter("size_label")
def size_label(value):
    """Whole numbers are container gallons (1, 5, 15...); text like QT or 6pk is kept."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{value:g} gal"
    return display_value(value)

@app.template_filter("money")
def money(value):
    try:
        return f"${float(value):,.2f}"
    except (TypeError, ValueError):
        return value

@app.route("/")
def index():
    q = request.args.get("q", "").strip().lower()
    headers, rows = headers_and_rows()
    total_all = len(rows)

    if q:
        filtered = []
        for row in rows:
            haystack = " ".join(display_value(row[h]) for h in headers).lower()
            if q in haystack:
                filtered.append(row)
        rows = filtered

    total = len(rows)
    pages = max(1, math.ceil(total / PAGE_SIZE))
    page = min(max(request.args.get("page", 1, type=int), 1), pages)
    start = (page - 1) * PAGE_SIZE
    page_rows = rows[start:start + PAGE_SIZE]

    return render_template("index.html", headers=headers, rows=page_rows,
                           q=request.args.get("q", ""), total=total, total_all=total_all,
                           page=page, pages=pages, first=start + 1 if total else 0,
                           last=start + len(page_rows),
                           check_field=CHECK_FIELD, has_check=CHECK_FIELD in headers,
                           today=date.today().strftime("%Y-%m-%d"), stock_options=stock_options())

def date_cell_format(ws, col):
    """Reuse the number format already used by dates in this column (e.g. m/d)."""
    for r in range(2, ws.max_row + 1):
        c = ws.cell(r, col)
        if isinstance(c.value, datetime):
            return c.number_format
    return "m/d;@"

@app.route("/check/<int:excel_row>", methods=["POST"])
def check(excel_row):
    """Tick/untick the 'Confirm' checkbox. Ticked -> LAST CHECK = today.
    Unticked -> LAST CHECK goes back to what it was before (sent by the page)."""
    data = request.get_json(silent=True) or {}
    checked = bool(data.get("checked"))
    previous = str(data.get("previous") or "").strip()
    try:
        with LOCK:
            wb, ws = get_workbook()
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            if CHECK_FIELD not in headers or not (2 <= excel_row <= ws.max_row):
                wb.close()
                return jsonify(ok=False, error="Record or column not found."), 404
            col = headers.index(CHECK_FIELD) + 1
            cell = ws.cell(excel_row, col)
            old = cell.value

            if checked:
                new = datetime.combine(date.today(), datetime.min.time())
            elif previous:
                new = datetime.strptime(previous, "%Y-%m-%d")
            else:
                new = None

            fmt = date_cell_format(ws, col)
            cell.value = new
            if new is not None:
                cell.number_format = fmt

            restore_stock_validation(ws, wb)
            wb.save(XLSX)
            wb.close()
    except PermissionError:
        return jsonify(ok=False, error="Could not save - is plant-index.xlsx open in Excel? Close it and try again."), 500
    except Exception as exc:
        return jsonify(ok=False, error=f"Could not save: {exc}"), 500

    return jsonify(ok=True,
                   date=display_value(new),
                   previous=display_value(old) if checked else "")

@app.route("/stock-note/<int:excel_row>", methods=["POST"])
def update_stock_note(excel_row):
    """Save a stock status directly from the inventory table."""
    data = request.get_json(silent=True) or {}
    value = str(data.get("value") or "").strip()
    allowed = stock_options()
    if value and value not in allowed:
        return jsonify(ok=False, error="Choose one of the available stock options."), 400
    try:
        with LOCK:
            wb, ws = get_workbook()
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            if "STOCK NOTE" not in headers or not (2 <= excel_row <= ws.max_row):
                wb.close()
                return jsonify(ok=False, error="Record or column not found."), 404
            ws.cell(excel_row, headers.index("STOCK NOTE") + 1).value = value or None
            restore_stock_validation(ws, wb)
            wb.save(XLSX)
            wb.close()
    except PermissionError:
        return jsonify(ok=False, error="Could not save - close plant-index.xlsx in Excel and try again."), 500
    except Exception as exc:
        return jsonify(ok=False, error=f"Could not save: {exc}"), 500
    return jsonify(ok=True, value=value)


@app.route("/inline-field/<int:excel_row>", methods=["POST"])
def update_inline_field(excel_row):
    data = request.get_json(silent=True) or {}
    field = str(data.get("field") or "")
    value = str(data.get("value") or "").strip()
    if field not in {"#", "ROW"}:
        return jsonify(ok=False, error="Field is not editable here."), 400
    try:
        with LOCK:
            wb, ws = get_workbook()
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            if field not in headers or not (2 <= excel_row <= ws.max_row):
                wb.close()
                return jsonify(ok=False, error="Record or column not found."), 404
            cell = ws.cell(excel_row, headers.index(field) + 1)
            cell.value = value or None
            wb.save(XLSX)
            wb.close()
    except PermissionError:
        return jsonify(ok=False, error="Close plant-index.xlsx in Excel and try again."), 500
    except Exception as exc:
        return jsonify(ok=False, error=f"Could not save: {exc}"), 500
    return jsonify(ok=True, value=value)

@app.route("/upload-photos", methods=["POST"])
def upload_photos():
    """Save image files dropped/browsed into the form to the local photos folder."""
    uploads = request.files.getlist("photos")
    if not uploads or all(not item.filename for item in uploads):
        return jsonify(ok=False, error="Choose or drop at least one image file."), 400

    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    saved = []
    errors = []
    for item in uploads:
        if not item or not item.filename:
            continue
        original = Path(item.filename.replace("\\", "/")).name
        safe_name = secure_filename(original)
        suffix = Path(safe_name).suffix.lower()
        if not safe_name or suffix not in IMAGE_EXTS:
            errors.append(f"Skipped {original or 'unnamed file'}: unsupported image type.")
            continue

        # Dropping a file already in the photos folder is common. If the existing
        # file is byte-for-byte identical, reuse it rather than making a duplicate.
        target = PHOTO_DIR / safe_name
        try:
            content = item.read()
            if target.exists() and target.read_bytes() == content:
                saved.append(target.relative_to(PHOTO_DIR).as_posix())
                continue
            # Never silently overwrite a different photo with the same name.
            if target.exists():
                stem = target.stem
                ext = target.suffix
                counter = 2
                while target.exists():
                    target = PHOTO_DIR / f"{stem}-{counter}{ext}"
                    counter += 1
            target.write_bytes(content)
            saved.append(target.relative_to(PHOTO_DIR).as_posix())
        except Exception as exc:
            errors.append(f"Could not save {original}: {exc}")

    if not saved:
        return jsonify(ok=False, error=" ".join(errors) or "No image files were saved."), 400
    return jsonify(ok=True, files=saved, warnings=errors)


@app.route("/photos/<path:filename>")
def photo(filename):
    # send_from_directory refuses paths that escape PHOTO_DIR.
    return send_from_directory(PHOTO_DIR, filename)

def size_sort_key(row):
    v = row.get("SIZE")
    return (0, float(v), "") if isinstance(v, (int, float)) else (1, 0, display_value(v).lower())

@app.route("/plant/<slug>")
def plant(slug):
    headers, rows = headers_and_rows()
    group = [r for r in rows if r["_slug"] == slug]
    if not group:
        abort(404)
    group.sort(key=size_sort_key)

    def first(field):
        for r in group:
            if r.get(field) not in (None, ""):
                return r[field]
        return None

    info = {h: first(h) for h in headers if h}
    image_paths = parse_image_paths(first(IMAGE_FIELD)) if IMAGE_FIELD in headers else []
    image_gallery = [{"path": rel, "exists": photo_exists(rel)} for rel in image_paths]
    image_rel = image_paths[0] if image_paths else ""
    if not image_paths:
        photo_state = "none"
    else:
        photo_state = "ok" if any(item["exists"] for item in image_gallery) else "missing"

    return render_template("plant.html", info=info, sizes=group, headers=headers,
                           staff_fields=[f for f in STAFF_FIELDS if f in headers],
                           show_staff=SHOW_STAFF_DETAILS,
                           image_rel=image_rel, image_gallery=image_gallery, photo_state=photo_state,
                           has_image_col=IMAGE_FIELD in headers,
                           back_q=request.args.get("q", ""), back_page=request.args.get("page", 1))

def form_context(headers, rows, row, mode, return_q, return_page):
    return dict(headers=headers, row=row, mode=mode, return_q=return_q,
                return_page=return_page, sections=form_sections(headers),
                wide=WIDE_FIELDS, suggest=suggestions(headers, rows),
                stock_options=stock_options(), photo_files=photo_files(),
                image_field=IMAGE_FIELD, has_image_col=IMAGE_FIELD in headers)

@app.route("/edit/<int:excel_row>", methods=["GET", "POST"])
def edit(excel_row):
    headers, rows = headers_and_rows()
    row = next((r for r in rows if r["_row"] == excel_row), None)
    if row is None:
        return "Record not found", 404

    if request.method == "POST":
        save_form(excel_row)
        flash("Record saved.")
        return redirect(url_for("index", q=request.form.get("_return_q", ""),
                                page=request.form.get("_return_page", 1)))

    return render_template("form.html", **form_context(
        headers, rows, row, "Edit", request.args.get("q", ""), request.args.get("page", 1)))

@app.route("/add", methods=["GET", "POST"])
def add():
    headers, rows = headers_and_rows()

    if request.method == "POST":
        save_form()
        flash("New record added.")
        return redirect(url_for("index"))

    row = {"_row": None}
    for h in headers:
        row[h] = None
    return render_template("form.html", **form_context(headers, rows, row, "Add", "", 1))


@app.route("/update-site", methods=["POST"])
def update_site():
    """Build the public, static website from the current admin workbook and photos."""
    builder = BASE / "build_site.py"
    try:
        result = subprocess.run(
            [sys.executable, str(builder)],
            cwd=str(BASE), capture_output=True, text=True, timeout=900
        )
    except subprocess.TimeoutExpired:
        flash("Site update took longer than 15 minutes and was stopped. Check the photos and try again.")
        return redirect(url_for("index"))
    except Exception as exc:
        flash(f"Could not start site update: {exc}")
        return redirect(url_for("index"))

    if result.returncode == 0:
        import shutil
        generated = BASE / "public-site"
        destination = BASE.parent / "teamlemuria.github.io"
        if destination.is_dir():
            try:
                shutil.copytree(generated, destination, dirs_exist_ok=True)
                flash("Website updated in public-site and copied to teamlemuria.github.io.")
            except Exception as exc:
                flash(f"Website built successfully in public-site; optional site-folder copy failed: {exc}")
        else:
            flash("Website updated successfully in public-site. No sibling publishing folder was found; local files are ready.")
    else:
        details = (result.stderr or result.stdout or "No error details were returned.").strip().splitlines()
        summary = "\n".join(details[-4:])
        log_path = BASE / "site-build-error.log"
        log_path.write_text((result.stdout or "") + "\n--- STDERR ---\n" + (result.stderr or ""), encoding="utf-8")
        flash(f"Site update failed. Details were saved to site-build-error.log. {summary[-500:]}")
    return redirect(url_for("index"))

if __name__ == "__main__":
    import webbrowser
    webbrowser.open("http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
