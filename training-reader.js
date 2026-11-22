// Focus Training on your phone: reading the book. The page-turning comes from Flint's reader
// (book pages laid out as columns, tap the sides or swipe to turn, Kindle-style themes, e-ink mode).
"use strict";

(() => {
  const T = window.T;
  const FONTS = {
    literata: ["Literata", '"Literata Variable", Georgia, serif', 1],
    sourceserif: ["Source Serif", '"Source Serif 4 Variable", Georgia, serif', 1],
    atkinson: ["Atkinson Hyperlegible", '"Atkinson Hyperlegible", system-ui, sans-serif', 1],
    georgia: ["Georgia", 'Georgia, "Times New Roman", serif', 1],
    system: ["System sans", 'system-ui, -apple-system, "Segoe UI", Roboto, sans-serif', 1],
  };
  const SIZES = [14, 15, 16, 17, 18, 19, 20, 22, 24, 26, 29, 32];
  const SPACING = { tight: 1.42, normal: 1.65, loose: 1.95 };
  const MARGINS = { narrow: 16, normal: 28, wide: 48 };
  const THEMES = ["auto", "white", "cream", "green", "dark", "black"];
  const DEFAULTS = { font: "literata", size: 4, spacing: "normal", margin: "normal", theme: "auto", layout: "pages", eink: "auto" };
  const P = (() => { try { return { ...DEFAULTS, ...JSON.parse(localStorage.getItem("focus-train-read") || "{}") }; } catch { return { ...DEFAULTS }; } })();
  const save = () => { try { localStorage.setItem("focus-train-read", JSON.stringify(P)); } catch { /* private mode */ } };
  const EINK_GUESS = /boox|onyx|e-?ink|kobo|tolino|bigme|hisense|pocketbook|meebook|likebook/i.test(navigator.userAgent)
    || matchMedia("(update: slow)").matches || matchMedia("(monochrome)").matches;
  const isEink = () => P.eink === "on" || (P.eink === "auto" && EINK_GUESS);
  const fsPx = () => Math.round(SIZES[P.size] * (FONTS[P.font] || FONTS.literata)[2]);
  const HL = { yellow: "hl-y", green: "hl-g", blue: "hl-b" };

  const R = { el: null, page: 0, total: 1, step: 0, opts: null, chrome: false };
  T.reader = { prefs: P, savePrefs: save, isEink, THEMES, FONTS, SIZES };

  // ---------- The book as HTML ----------

  function runsHtml(p) {
    return (p.Runs || []).map(r => {
      let t = T.esc(r.Text || "");
      if (r.Italic) t = `<i>${t}</i>`;
      if (r.Bold) t = `<b>${t}</b>`;
      if (r.Super) t = `<sup>${t}</sup>`;
      return t;
    }).join("");
  }
  function chartHtml(c) {
    const head = "<tr><th></th>" + (c.SeriesNames || []).map(n => `<th>${T.esc(n)}</th>`).join("") + "</tr>";
    const rows = (c.Categories || []).map((cat, i) => `<tr><th>${T.esc(cat)}</th>` + (c.SeriesValues || []).map(s => `<td>${T.esc(s[i] ?? "")}</td>`).join("") + "</tr>").join("");
    return `<div class="tr-tablewrap"><table class="tr-chart"><caption>${T.esc(c.Title || "Chart")}</caption>${head}${rows}</table></div>`;
  }
  /** Paragraphs from..to of a book, ready to read (each block knows its paragraph number). */
  T.bookHtml = (courseId, from, to) => {
    const book = T.B(courseId).book;
    const out = [];
    let i = Math.max(0, from);
    to = Math.min(to, book.length);
    while (i < to) {
      const p = book[i];
      if (p.InTable) {
        const id = p.TableId, rows = {};
        const start = i;
        while (i < to && book[i].InTable && book[i].TableId === id) { const q = book[i]; ((rows[q.Row] ||= {})[q.Col] ||= []).push(`<span data-para="${i}">${runsHtml(q)}</span>`); i++; }
        const body = Object.keys(rows).sort((a, b) => a - b).map(r => "<tr>" + Object.keys(rows[r]).sort((a, b) => a - b).map(c => `<td>${rows[r][c].join("<br>")}</td>`).join("") + "</tr>").join("");
        out.push(`<div class="tr-tablewrap" data-para="${start}"><table>${body}</table></div>`);
        continue;
      }
      if (p.Sidebar) {
        const start = i, parts = [];
        while (i < to && book[i].Sidebar && !book[i].InTable) { parts.push(blockHtml(book[i], i)); i++; }
        out.push(`<aside class="tr-side" data-para="${start}">${parts.join("")}</aside>`);
        continue;
      }
      out.push(blockHtml(p, i));
      i++;
    }
    return out.join("");
  };
  function blockHtml(p, i) {
    if (p.Image) return `<figure data-para="${i}"><img alt="" data-img="${T.esc(p.Image)}"><figcaption class="tr-missing">Picture: it downloads the next time you sync on Wi-Fi.</figcaption></figure>`;
    if (p.Chart) return `<div data-para="${i}">${chartHtml(p.Chart)}</div>`;
    const text = runsHtml(p);
    if (!text.trim()) return "";
    if (p.Heading > 0) { const h = Math.min(4, p.Heading + 1); return `<h${h} data-para="${i}">${text}</h${h}>`; }
    if (p.Caption) return `<p class="tr-cap" data-para="${i}">${text}</p>`;
    if (p.ListMark) return `<p class="tr-li" style="--lvl:${p.ListLevel || 0}" data-para="${i}"><span class="tr-mark">${T.esc(p.ListMark)}</span>${text}</p>`;
    return `<p data-para="${i}">${text}</p>`;
  }
  /** Puts pictures (kept on this device) into the page. */
  T.fillImages = async (root, courseId) => {
    for (const img of root.querySelectorAll("img[data-img]")) {
      const url = await T.mediaUrl(courseId, img.dataset.img);
      if (url) { img.src = url; img.nextElementSibling?.remove(); }
      else img.remove();
    }
  };
  /** Marks your highlights in the text. */
  T.markHighlights = (root, courseId) => {
    for (const h of T.highlights(courseId)) {
      const host = root.querySelector(`[data-para="${h.para}"]`);
      if (!host || !h.quote) continue;
      wrapText(host, h.quote, h);
    }
  };
  function wrapText(host, quote, h) {
    const want = quote.trim();
    const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT);
    const nodes = []; let all = "";
    while (walker.nextNode()) { nodes.push([walker.currentNode, all.length]); all += walker.currentNode.nodeValue; }
    const at = all.indexOf(want);
    if (at < 0) return;
    const end = at + want.length;
    for (const [node, off] of nodes) {
      const a = Math.max(at, off), b = Math.min(end, off + node.nodeValue.length);
      if (a >= b) continue;
      const range = document.createRange();
      range.setStart(node, a - off); range.setEnd(node, b - off);
      const mark = document.createElement("mark");
      mark.className = HL[h.color] || HL.yellow; mark.dataset.hl = h.id;
      range.surroundContents(mark);
    }
  }

  // ---------- The reader ----------

  /**
   * Opens the book for reading.
   * opts: { courseId, from, to, title, chapter, goTo (a paragraph), doneLabel, onDone, onClose }
   */
  T.openReader = opts => {
    closeReader();
    R.opts = opts;
    const el = document.createElement("div");
    el.className = "trr"; el.setAttribute("role", "dialog"); el.setAttribute("aria-label", opts.title || "Book");
    el.innerHTML = `
      <div class="trr-top">
        <button class="trr-ic" data-act="close" aria-label="Close">✕</button>
        <div class="trr-title">${T.esc(opts.title || "")}</div>
        <button class="trr-ic" data-act="aa" aria-label="Text and page">Aa</button>
      </div>
      <div class="trr-frame"><div class="trr-clip"><div class="trr-cols tr-prose">${T.bookHtml(opts.courseId, opts.from, opts.to)}<div class="trr-end"></div></div></div></div>
      <div class="trr-foot"><span class="trr-prog"></span>${opts.onDone ? `<button class="tr-btn primary trr-done">${T.esc(opts.doneLabel || "Done reading")}</button>` : ""}</div>
      <div class="trr-pop" hidden></div>
      <div class="trr-aa" hidden></div>`;
    document.body.append(el);
    R.el = el;
    document.documentElement.classList.add("tr-reading");
    applyPrefs();
    const cols = el.querySelector(".trr-cols");
    T.markHighlights(cols, opts.courseId);
    T.fillImages(cols, opts.courseId).then(() => relayout());
    bind();
    requestAnimationFrame(() => {
      layout();
      if (opts.goTo != null) showPara(opts.goTo); else goPage(0, false);
    });
  };
  function closeReader() {
    if (!R.el) return;
    R.el.remove(); R.el = null;
    document.documentElement.classList.remove("tr-reading");
    window.removeEventListener("resize", relayout);
    document.removeEventListener("keydown", onKey);
    document.removeEventListener("selectionchange", onSelection);
  }
  T.closeReader = () => { const o = R.opts; closeReader(); o && o.onClose && o.onClose(); };
  /** Closing from the reader's own ✕ (or Done): through the phone's Back history when the page keeps one. */
  const requestClose = () => { const o = R.opts; if (o && o.requestClose) o.requestClose(); else T.closeReader(); };
  const q = sel => R.el.querySelector(sel);

  function applyPrefs() {
    if (!R.el) return;
    const f = FONTS[P.font] || FONTS.literata;
    const s = R.el.style;
    s.setProperty("--read", f[1]);
    s.setProperty("--fs", fsPx() + "px");
    s.setProperty("--lh", SPACING[P.spacing]);
    s.setProperty("--m", MARGINS[P.margin] + "px");
    R.el.dataset.rtheme = isEink() && P.theme === "auto" ? "white" : P.theme;
    R.el.dataset.layout = P.layout;
    document.documentElement.classList.toggle("eink", isEink());
  }

  function layout() {
    const fr = q(".trr-frame"), clip = q(".trr-clip"), cols = q(".trr-cols");
    if (P.layout === "scroll") { cols.removeAttribute("style"); clip.removeAttribute("style"); R.total = 1; updateProgress(); return; }
    fr.scrollTop = 0;
    const W = fr.clientWidth, H = fr.clientHeight, m = MARGINS[P.margin];
    const two = W >= 880 && W / H > 1.15;
    const gap = two ? Math.round(Math.max(48, m * 1.6)) : Math.max(40, m);
    const maxCol = fsPx() * 34;
    const cw = Math.floor(two ? Math.min((W - 2 * m - gap) / 2, maxCol) : Math.min(W - 2 * m, maxCol));
    const box = two ? 2 * cw + gap : cw;
    Object.assign(clip.style, { position: "absolute", top: "0", bottom: "0", left: `${Math.floor((W - box) / 2)}px`, width: `${box}px`, overflow: "hidden" });
    Object.assign(cols.style, { width: `${box}px`, height: `${H}px`, columnWidth: `${cw}px`, columnGap: `${gap}px` });
    R.step = box + gap;
    R.total = Math.max(1, Math.ceil((cols.scrollWidth + gap) / R.step - 0.02));
  }
  function relayout() {
    if (!R.el) return;
    const frac = R.total > 1 ? R.page / (R.total - 1) : 0;
    layout();
    goPage(Math.round(frac * (R.total - 1)), false);
  }
  function goPage(p, animate = true) {
    if (P.layout === "scroll") { updateProgress(); return; }
    R.page = Math.max(0, Math.min(R.total - 1, p));
    const cols = q(".trr-cols");
    const anim = animate && !isEink();
    if (!anim) cols.style.transition = "none";
    cols.style.transform = `translateX(${-R.page * R.step}px)`;
    if (!anim) { void cols.offsetWidth; cols.style.transition = ""; }
    updateProgress();
  }
  function turn(d) {
    hidePop();
    if (P.layout !== "pages") { const fr = q(".trr-frame"); fr.scrollBy({ top: d * fr.clientHeight * 0.9, behavior: isEink() ? "auto" : "smooth" }); return; }
    goPage(R.page + d);
    setChrome(false);
  }
  /** Turns to the page with a paragraph on it (and flashes it). */
  function showPara(i) {
    const cols = q(".trr-cols");
    let el = cols.querySelector(`[data-para="${i}"]`);
    if (!el) { const all = [...cols.querySelectorAll("[data-para]")]; el = all.filter(x => +x.dataset.para <= i).pop() || all[0]; }
    if (!el) return goPage(0, false);
    if (P.layout === "scroll") el.scrollIntoView({ block: "center" });
    else {
      const x = el.getBoundingClientRect().left - cols.getBoundingClientRect().left;
      goPage(Math.floor((x + 2) / R.step), false);
    }
    el.classList.add("tr-flash");
    setTimeout(() => el.classList.remove("tr-flash"), 2400);
  }
  function updateProgress() {
    if (!R.el) return;
    const fr = q(".trr-frame");
    let atEnd;
    if (P.layout === "scroll") {
      const h = fr.scrollHeight - fr.clientHeight;
      const frac = h > 0 ? fr.scrollTop / h : 1;
      q(".trr-prog").textContent = `${Math.round(frac * 100)}%`;
      atEnd = frac > 0.97;
    } else {
      q(".trr-prog").textContent = `Page ${R.page + 1} of ${R.total}`;
      atEnd = R.page >= R.total - 1;
    }
    const done = q(".trr-done");
    if (done) done.classList.toggle("ready", atEnd);
  }
  function setChrome(on) { R.chrome = on; R.el.classList.toggle("chrome", on); }

  function bind() {
    const fr = q(".trr-frame");
    let sx = 0, sy = 0, swiped = 0;
    fr.addEventListener("touchstart", e => { sx = e.touches[0].clientX; sy = e.touches[0].clientY; }, { passive: true });
    fr.addEventListener("touchend", e => {
      if (!getSelection().isCollapsed || P.layout !== "pages") return;
      const t = e.changedTouches[0], dx = t.clientX - sx, dy = t.clientY - sy;
      if (Math.abs(dx) > 45 && Math.abs(dx) > Math.abs(dy) * 1.4) { swiped = Date.now(); turn(dx < 0 ? 1 : -1); }
    }, { passive: true });
    fr.addEventListener("click", e => {
      if (Date.now() - swiped < 400 || !getSelection().isCollapsed) return;
      const mark = e.target.closest("mark[data-hl]");
      if (mark) { showMarkPop(mark); return; }
      if (!q(".trr-pop").hidden) { hidePop(); return; }
      if (!q(".trr-aa").hidden) { q(".trr-aa").hidden = true; return; }
      const x = (e.clientX - fr.getBoundingClientRect().left) / fr.clientWidth;
      if (P.layout === "pages" && x < 0.28) turn(-1);
      else if (P.layout === "pages" && x > 0.72) turn(1);
      else setChrome(!R.chrome);
    });
    fr.addEventListener("scroll", () => { if (P.layout === "scroll") updateProgress(); }, { passive: true });
    R.el.querySelector("[data-act=close]").onclick = () => requestClose();
    R.el.querySelector("[data-act=aa]").onclick = () => toggleAa();
    const done = q(".trr-done");
    if (done) done.onclick = () => { const o = R.opts; requestClose(); o.onDone(); };
    q(".trr-pop").addEventListener("pointerdown", e => e.preventDefault());
    window.addEventListener("resize", relayout);
    document.addEventListener("keydown", onKey);
    document.addEventListener("selectionchange", onSelection);
  }
  function onKey(e) {
    if (!R.el || /INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName)) return;
    if (["ArrowRight", "ArrowDown", "PageDown", " "].includes(e.key)) { e.preventDefault(); turn(1); }
    else if (["ArrowLeft", "ArrowUp", "PageUp"].includes(e.key)) { e.preventDefault(); turn(-1); }
    else if (e.key === "Escape") requestClose();
  }

  // ---------- Highlights (select text, pick a color) ----------

  let selTimer = null;
  function onSelection() {
    clearTimeout(selTimer);
    selTimer = setTimeout(() => {
      if (!R.el || R.opts.chapter == null) return;
      const sel = getSelection();
      if (sel.isCollapsed || !sel.rangeCount) return;
      const range = sel.getRangeAt(0);
      const cols = q(".trr-cols");
      if (!cols.contains(range.commonAncestorContainer)) return;
      const host = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement).closest("[data-para]");
      const quote = sel.toString().trim().replace(/\s+/g, " ");
      if (!host || quote.length < 3) return;
      const r = range.getBoundingClientRect();
      showPop(r, `<button data-c="yellow" class="sw hl-y" aria-label="Yellow"></button><button data-c="green" class="sw hl-g" aria-label="Green"></button><button data-c="blue" class="sw hl-b" aria-label="Blue"></button>`, b => {
        const color = b.dataset.c;
        if (!color) return;
        const para = +host.dataset.para;
        const id = Math.random().toString(16).slice(2, 14).padEnd(12, "0");
        T.add({ Kind: "highlight", Question: id, Chapter: R.opts.chapter, Para: para, Quote: quote.slice(0, 600), Color: color }, R.opts.courseId);
        getSelection().removeAllRanges();
        wrapText(host, quote, { id, color });
        hidePop();
      });
    }, 350);
  }
  function showMarkPop(mark) {
    showPop(mark.getBoundingClientRect(), `<button data-x="1" class="tr-btn">Remove highlight</button>`, () => {
      T.add({ Kind: "unhighlight", Question: mark.dataset.hl }, R.opts.courseId);
      R.el.querySelectorAll(`mark[data-hl="${CSS.escape(mark.dataset.hl)}"]`).forEach(m => m.replaceWith(...m.childNodes));
      hidePop();
    });
  }
  function showPop(rect, html, onPick) {
    const pop = q(".trr-pop");
    pop.innerHTML = html; pop.hidden = false;
    const w = pop.offsetWidth, h = pop.offsetHeight;
    const top = rect.top > h + 70 ? rect.top - h - 10 : rect.bottom + 10;
    pop.style.top = Math.max(8, Math.min(innerHeight - h - 8, top)) + "px";
    pop.style.left = Math.max(8, Math.min(innerWidth - w - 8, rect.left + rect.width / 2 - w / 2)) + "px";
    pop.onclick = e => { const b = e.target.closest("button"); if (b) onPick(b); };
  }
  function hidePop() { if (R.el) q(".trr-pop").hidden = true; }

  // ---------- Text and page settings ----------

  function toggleAa() {
    const aa = q(".trr-aa");
    if (!aa.hidden) { aa.hidden = true; return; }
    aa.hidden = false;
    paintAa();
  }
  function seg(key, list, label) {
    return `<div class="trr-row"><span>${label}</span><div class="tr-seg">${list.map(([v, t]) => `<button data-k="${key}" data-v="${v}" aria-pressed="${String(P[key]) === String(v)}">${t}</button>`).join("")}</div></div>`;
  }
  function paintAa() {
    const aa = q(".trr-aa");
    aa.innerHTML = `
      <div class="trr-row"><span>Size</span><div class="tr-seg"><button data-size="-1" aria-label="Smaller">A−</button><b>${fsPx()}</b><button data-size="1" aria-label="Bigger">A+</button></div></div>
      ${seg("font", Object.entries(FONTS).map(([k, f]) => [k, f[0]]), "Font")}
      ${seg("theme", THEMES.map(t => [t, t[0].toUpperCase() + t.slice(1)]), "Page")}
      ${seg("spacing", [["tight", "Tight"], ["normal", "Normal"], ["loose", "Loose"]], "Lines")}
      ${seg("margin", [["narrow", "Narrow"], ["normal", "Normal"], ["wide", "Wide"]], "Margins")}
      ${seg("layout", [["pages", "Pages"], ["scroll", "Scroll"]], "Turn")}
      ${seg("eink", [["auto", "Auto"], ["on", "On"], ["off", "Off"]], "E-ink")}`;
    aa.onclick = e => {
      const b = e.target.closest("button");
      if (!b) return;
      if (b.dataset.size) P.size = Math.max(0, Math.min(SIZES.length - 1, P.size + +b.dataset.size));
      else if (b.dataset.k) P[b.dataset.k] = b.dataset.v;
      save(); applyPrefs(); paintAa(); relayout();
    };
  }
})();
