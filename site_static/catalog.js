// Plant list: filters the data in plants-data.js in the browser (no server needed).
(function () {
  var PAGE = 48;
  var all = window.PLANTS || [];
  var $ = function (id) { return document.getElementById(id); };
  var q = $("q"), category = $("category"), sun = $("sun"), trait = $("trait"), instock = $("instock");
  var results = $("results"), count = $("count"), more = $("more");
  var matches = [], shown = 0;

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function card(p) {
    var img = p.i || "static/placeholder.svg";
    return '<a class="card" href="plant/' + esc(p.s) + '.html">' +
      '<img src="' + esc(img) + '" alt="" loading="lazy">' +
      '<div class="card-body"><h3>' + esc(p.n) + '</h3><p class="botanical"><i>' + esc(p.b) + '</i></p>' +
      '<p class="card-meta"><span>' + (p.p ? "From " + esc(p.p) : "") + '</span>' +
      '<span class="badge ' + esc(p.k) + '">' + esc(p.st) + '</span></p></div></a>';
  }

  function showMore() {
    var next = matches.slice(shown, shown + PAGE);
    results.insertAdjacentHTML("beforeend", next.map(card).join(""));
    shown += next.length;
    more.hidden = shown >= matches.length;
    count.textContent = matches.length
      ? "Showing " + shown + " of " + matches.length + " plants"
      : "No plants match. Try fewer filters or a different word.";
  }

  function apply() {
    var words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
    matches = all.filter(function (p) {
      if (category.value && p.c !== category.value) return false;
      if (sun.value && p.u.indexOf(sun.value) < 0) return false;
      if (trait.value && p.t.indexOf(trait.value) < 0) return false;
      if (instock.checked && p.k !== "ok") return false;
      return words.every(function (w) { return p.h.indexOf(w) >= 0; });
    });
    results.innerHTML = "";
    shown = 0;
    showMore();
    var params = new URLSearchParams();
    if (q.value) params.set("q", q.value);
    if (category.value) params.set("category", category.value);
    if (sun.value) params.set("sun", sun.value);
    if (trait.value) params.set("trait", trait.value);
    if (instock.checked) params.set("instock", "1");
    var qs = params.toString();
    history.replaceState(null, "", qs ? "?" + qs : location.pathname);
  }

  var p = new URLSearchParams(location.search);
  q.value = p.get("q") || "";
  category.value = p.get("category") || "";
  sun.value = p.get("sun") || "";
  trait.value = p.get("trait") || "";
  instock.checked = p.get("instock") === "1";

  [q, category, sun, trait, instock].forEach(function (el) {
    el.addEventListener(el === q ? "input" : "change", apply);
  });
  more.addEventListener("click", showMore);
  apply();
})();
