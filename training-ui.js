// Focus Training on your phone: the screens (dashboard, lessons, practice), made for touch.
// Everything you do is saved the way Focus on your computers saves it, so a lesson finished here
// is finished there too.
"use strict";

(() => {
  const T = window.T;
  const E = T.esc;
  const UI = (() => { try { return { books: 10, cards: 20, timeline: 6, greek: 6, ...JSON.parse(localStorage.getItem("focus-train-ui") || "{}") }; } catch { return { books: 10, cards: 20, timeline: 6, greek: 6 }; } })();
  const saveUI = () => { try { localStorage.setItem("focus-train-ui", JSON.stringify(UI)); } catch { /* private mode */ } };
  const $ = id => document.getElementById(id);
  const note = text => (typeof toast === "function" ? toast(text) : console.log(text));
  const plural = (n, w, ws) => `${n} ${n === 1 ? w : ws || w + "s"}`;

  // ---------- Layers (the phone's Back button closes the top one) ----------

  const stack = [];
  function openLayer(close) { stack.push(close); history.pushState({ tr: stack.length }, ""); }
  const back = () => history.back();
  window.addEventListener("popstate", () => { const close = stack.pop(); if (close) close(); });

  // ---------- The Training screen ----------

  let root = null, started = false;
  function ensureRoot() {
    if (root) return root;
    root = document.createElement("section");
    root.id = "training"; root.className = "tr hidden";
    root.innerHTML = `
      <div class="tr-wrap">
        <header class="tr-head">
          <button class="tr-ic" data-act="tasks" aria-label="Back to tasks">‹</button>
          <h1>Training</h1>
          <span class="tr-sync" id="trSync" role="status"></span>
          <button class="tr-ic" data-act="settings" aria-label="Training settings"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="3"/><path d="M12 3v2M12 19v2M3 12h2M19 12h2M5.6 5.6 7 7M17 17l1.4 1.4M5.6 18.4 7 17M17 7l1.4-1.4"/></svg></button>
        </header>
        <div id="trDash"></div>
      </div>`;
    document.body.append(root);
    root.querySelector("[data-act=tasks]").onclick = back;
    root.querySelector("[data-act=settings]").onclick = openSettings;
    root.addEventListener("click", onDashClick);
    T.onChange(() => { if (!root.classList.contains("hidden")) paintDash(); paintSync(); });
    return root;
  }
  function openTraining() {
    ensureRoot();
    if (!started) { started = true; localStorage.setItem("focus-train-used", "1"); T.start(); }
    else T.sync();
    root.classList.remove("hidden");
    document.body.classList.add("tr-open");
    document.documentElement.classList.toggle("eink", T.reader.isEink());
    paintDash();
    window.scrollTo(0, 0);
    openLayer(() => { root.classList.add("hidden"); document.body.classList.remove("tr-open"); document.documentElement.classList.remove("eink"); });
  }
  T.openTraining = openTraining;

  function paintSync() {
    const el = $("trSync");
    if (!el) return;
    const s = T.status;
    el.className = "tr-sync " + (s.busy ? "busy" : s.error ? "off" : s.offline ? "off" : s.pending ? "wait" : "ok");
    el.textContent = s.busy ? "Syncing…" : s.error ? "Sync problem" : s.offline ? "Offline" : s.pending ? "Waiting to save" : s.lastSync ? "Synced" : "";
    el.title = s.text || "";
  }

  // ---------- The dashboard ----------

  const courses = () => T.config.Courses || [];
  const ids = () => courses().map(c => c.Id);
  function weekStats() {
    const now = new Date(), mon = T.addDays(T.startOfDay(now), -((now.getDay() + 6) % 7)), from = T.ymd(mon);
    const lessons = new Set(T.events.filter(e => e.Kind === "lessonDone" && e.Lesson >= 0 && e.Day >= from).map(e => T.courseOf(e) + "|" + e.Lesson)).size;
    const answers = T.events.filter(e => e.Kind === "answer" && e.Day >= from);
    const right = answers.filter(e => e.Grade >= 1 || (e.Grade < 0 && e.Correct && e.Sure)).length;
    let planned = 0;
    for (let d = mon; d <= T.addDays(mon, 6); d = T.addDays(d, 1)) planned += T.booksFor(d).length;
    return { lessons, planned, answers: answers.length, right };
  }
  const dueBooks = () => ids().reduce((n, id) => n + T.dueReview(id).filter(x => x.q.kind !== "card").length, 0);
  const dueCards = () => ids().reduce((n, id) => n + T.dueReview(id).filter(x => x.q.kind === "card").length, 0);

  function paintDash() {
    const dash = $("trDash");
    if (!dash) return;
    if (!courses().length) {
      const st = T.status;
      dash.innerHTML = st.busy
        ? `<div class="tr-empty"><p><b>Getting your training…</b></p><p class="tr-stepline">${E(st.step || "Connecting to OneDrive…")}</p><p>The first time, your books and lessons download from OneDrive. After that they're kept here for offline use.</p></div>`
        : st.error || st.offline
          ? `<div class="tr-empty"><p><b>${st.offline ? "You're offline" : "Couldn't get your training"}</b></p><p class="tr-errline">${E(st.offline ? "Connect to the internet and try again." : st.error)}</p><button class="tr-btn primary" data-act="sync">Try again</button></div>`
          : st.lastSync
            ? `<div class="tr-empty"><p><b>No books yet</b></p><p>Add a book on your computer (Training, Library). It shows up here after it syncs.</p><button class="tr-btn" data-act="sync">Sync now</button></div>`
            : `<div class="tr-empty"><p><b>Getting your training…</b></p><p class="tr-stepline">Starting…</p><button class="tr-btn" data-act="sync">Sync now</button></div>`;
      return;
    }
    const today = new Date(), off = T.dayOff(today);
    const parts = [];
    // Today
    const day = T.dayLessons(today);
    let todayHtml = "";
    if (off === "off") todayHtml = `<div class="tr-banner"><b>Training is off today.</b> No lessons; nothing counts as missed.</div>`;
    else {
      if (off === "books") todayHtml += `<div class="tr-banner"><b>Books off today.</b> Practice is still on.</div>`;
      if (!day.length && off !== "books") todayHtml += `<p class="tr-muted">No lesson planned today.</p>`;
      for (const { book, done } of day) todayHtml += lessonCard(book, done);
      const miss = T.missedLesson();
      if (miss && !off) todayHtml += `<div class="tr-row-card"><div><b>Catch up</b><small>${E(miss.book.Title)}: ${E(miss.day.toLocaleDateString([], { weekday: "long" }))}'s lesson wasn't done.</small></div><button class="tr-btn" data-act="lesson" data-book="${E(miss.book.Id)}">Start</button></div>`;
    }
    parts.push(`<h2 class="tr-h">Today · ${E(today.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" }))}</h2>${todayHtml}`);

    // Practice tiles
    const tiles = [];
    const tile = (act, title, count, sub) => tiles.push(`<button class="tr-tile" data-act="${act}"><span class="tr-tile-n${count ? "" : " zero"}">${count == null ? "·" : count}</span><b>${title}</b><small>${sub}</small></button>`);
    if (off !== "off") {
      tile("books", "Books practice", dueBooks(), "Questions due");
      tile("cards", "Flashcards", dueCards(), "Memory cards due");
      tile("timeline", "Timeline", T.timelineDue(ids()), "Dates to place");
      if (T.greek.vocab.length) tile("greek:vocab", "Greek words", T.greekDue("vocab"), "Words due");
      if (T.greek.alphabet.length) tile("greek:alphabet", "Alphabet", T.greekDue("alphabet"), "Letters due");
      if (T.greek.concepts.length) tile("greek:concepts", "Greek concepts", T.greekDue("concepts"), "Concepts due");
      if (T.greek.quizzes.length) tile("quiz", "Concept quiz", null, plural(T.greek.quizzes.length, "lesson"));
      if (courses().some(c => T.finished(c.Id))) tile("allbooks", "All books", null, "Finished books, mixed");
      parts.push(`<h2 class="tr-h">Practice</h2><div class="tr-tiles">${tiles.join("")}</div>`);
    }

    // Library
    const lib = [];
    for (const c of courses()) {
      const pct = Math.round(T.bookFraction(c.Id) * 100), next = T.nextLesson(c.Id);
      const sub = T.finished(c.Id) ? "Finished" : next ? `Next: ch. ${next.pack.chapter} · ${E(next.lesson.title || "Lesson " + next.number)}` : T.B(c.Id).lessons.length ? "Waiting for the next chapter pack" : "No chapter packs yet";
      lib.push(`<button class="tr-book" data-act="book" data-book="${E(c.Id)}"><span class="tr-cover" data-cover="${E(c.Id)}">${E((c.Title || "?").slice(0, 1))}</span><span class="tr-book-t"><b>${E(c.Title)}</b><small>${sub}</small><span class="tr-bar"><span style="width:${pct}%"></span></span></span><span class="tr-pct">${pct}%</span></button>`);
    }
    parts.push(`<h2 class="tr-h">Library</h2><div class="tr-lib">${lib.join("")}</div>`);

    // This week
    const w = weekStats();
    parts.push(`<h2 class="tr-h">This week</h2><div class="tr-stats"><div><b>${w.lessons}</b><small>${w.lessons === 1 ? "lesson" : "lessons"} done · ${w.planned} planned</small></div><div><b>${w.answers}</b><small>answers</small></div><div><b>${w.answers ? Math.round((100 * w.right) / w.answers) : 0}%</b><small>right and sure</small></div></div>`);
    const st = T.status;
    parts.push(`<p class="tr-foot">${E(st.busy ? st.step || "Syncing…" : st.text || "")}${st.media ? `<br>${E(st.media)}` : ""}${st.error ? `<br><button class="tr-link" data-act="sync">Try again</button>` : ""}</p>`);
    if (st.error) parts.unshift(`<div class="tr-banner">${E(st.error)}</div>`);
    dash.innerHTML = parts.join("");
    dash.querySelectorAll("[data-cover]").forEach(async el => { const c = T.course(el.dataset.cover); const url = c && await T.coverUrl(c); if (url) { el.textContent = ""; el.style.backgroundImage = `url("${url}")`; } });
    paintSync();
  }
  function lessonCard(book, done) {
    const exam = T.examDue(book.Id), next = T.nextLesson(book.Id);
    let title, sub, btn;
    if (done) { title = book.Title; sub = "Today's lesson is done. Nicely done."; btn = next ? `<button class="tr-btn" data-act="lesson" data-book="${E(book.Id)}">Extra lesson</button>` : ""; }
    else if (exam) { title = exam.title; sub = `${E(book.Title)} · ${exam.questions.length} questions, pass with 70%`; btn = `<button class="tr-btn primary" data-act="exam" data-book="${E(book.Id)}">Start exam</button>`; }
    else if (next) {
      const going = T.inProgress(next);
      title = next.lesson.title || `Lesson ${next.number}`;
      sub = `${E(book.Title)} · chapter ${next.pack.chapter}, lesson ${next.number} of ${next.pack.lessons.length}`;
      btn = `<button class="tr-btn primary" data-act="lesson" data-book="${E(book.Id)}">${going ? "Continue" : "Start"}</button>`;
    } else { title = book.Title; sub = T.finished(book.Id) ? "Finished. Its questions are in All books." : "Waiting for the next chapter pack."; btn = ""; }
    return `<div class="tr-lesson${done ? " done" : ""}"><small class="tr-role">${E(book.Role || "Today's book")}</small><b>${E(title)}</b><small>${sub}</small>${btn ? `<div class="tr-lesson-btns">${btn}</div>` : ""}</div>`;
  }
  function onDashClick(e) {
    const b = e.target.closest("[data-act]");
    if (!b || !root.contains(b) || b.closest(".trp, .trs")) return;
    const act = b.dataset.act, book = b.dataset.book;
    if (act === "sync") { T.sync(); paintDash(); }
    else if (act === "lesson") { const lr = T.nextLesson(book); if (lr) runLesson(lr); else note("No lesson ready in this book yet."); }
    else if (act === "exam") runExam(book);
    else if (act === "books") runBooksPractice();
    else if (act === "cards") runFlashcards();
    else if (act === "timeline") runTimelinePractice();
    else if (act.startsWith("greek:")) runGreek(act.slice(6));
    else if (act === "quiz") openQuizList();
    else if (act === "allbooks") runAllBooks();
    else if (act === "book") openBook(book);
  }

  // ---------- A full-screen player for lessons and practice ----------

  /** steps: [{ key, run(host, next) }]; onStep(key) saves a finished step (lessons). */
  function player(title, steps, { onStep, onClose } = {}) {
    const el = document.createElement("div");
    el.className = "trp";
    el.innerHTML = `<div class="trp-top"><button class="tr-ic" data-x aria-label="Close">✕</button><div class="trp-title">${E(title)}</div><span class="trp-count"></span></div><div class="trp-bar"><span></span></div><div class="trp-body"><div class="trp-page"></div></div>`;
    document.body.append(el);
    document.body.classList.add("trp-open");
    const page = el.querySelector(".trp-page"), body = el.querySelector(".trp-body");
    let i = -1, closed = false;
    const close = () => { if (closed) return; closed = true; el.remove(); if (!document.querySelector(".trp")) document.body.classList.remove("trp-open"); onClose && onClose(); paintDash(); };
    openLayer(close);
    el.querySelector("[data-x]").onclick = back;
    const go = () => {
      if (closed) return;
      if (i >= 0 && steps[i] && steps[i].key && onStep) onStep(steps[i].key);
      i++;
      if (i >= steps.length) { back(); return; }
      el.querySelector(".trp-count").textContent = steps.length > 2 ? `${i + 1} / ${steps.length}` : "";
      el.querySelector(".trp-bar span").style.width = `${Math.round((100 * i) / Math.max(1, steps.length - 1))}%`;
      page.innerHTML = ""; page.className = "trp-page";
      body.scrollTop = 0;
      steps[i].run(page, go);
    };
    go();
    return { close: back, el };
  }
  const btn = (label, cls = "") => `<button class="tr-btn ${cls}">${label}</button>`;
  /** A "Next" button that also works with Enter. */
  function nextButton(host, next, label = "Next") {
    const row = document.createElement("div"); row.className = "tr-act";
    row.innerHTML = btn(label + " →", "primary");
    host.append(row);
    const b = row.querySelector("button");
    b.onclick = () => { document.removeEventListener("keydown", onKey); next(); };
    const onKey = e => { if (e.key === "Enter" && document.body.contains(b) && !/INPUT|TEXTAREA/.test(document.activeElement?.tagName)) { e.preventDefault(); b.click(); } };
    document.addEventListener("keydown", onKey);
    b.focus({ preventScroll: true });
    return b;
  }
  const daysText = d => { const n = Math.round((T.startOfDay(d) - T.startOfDay(new Date())) / 86400000); return n <= 0 ? "today" : n === 1 ? "tomorrow" : `in ${n} days`; };

  // ----- Show in book -----
  function showInBook(lr, ref) {
    const para = T.locate(lr, ref);
    if (para < 0) { note("Couldn't find that spot in the book."); return; }
    const [cs, ce] = T.chapterRange(lr.course, lr.pack);
    T.openReader({ courseId: lr.course, from: cs, to: ce, title: `Chapter ${lr.pack.chapter}`, chapter: lr.pack.chapter, goTo: para, requestClose: back });
    openLayer(T.closeReader);
  }
  function qMenu(host, q, courseId) {
    const m = document.createElement("details"); m.className = "tr-more";
    m.innerHTML = `<summary aria-label="More">⋯</summary><div><button data-k="flag">Flag this question</button><button data-k="hide">Hide this question</button></div>`;
    m.onclick = e => {
      const b = e.target.closest("button[data-k]");
      if (!b) return;
      T.add({ Kind: b.dataset.k, Question: q.id, Concept: q.concept || "" }, courseId);
      b.disabled = true; b.textContent = b.dataset.k === "flag" ? "Flagged" : "Hidden (won't come up again)";
    };
    host.append(m);
  }

  // ----- Multiple choice -----
  function mcStep(host, lr, q, label, scored, next, tally) {
    host.innerHTML = `<div class="tr-step">${E(label)}</div><h2 class="tr-q">${E(q.prompt)}</h2>
      <div class="tr-opts">${(q.options || []).map((t, i) => `<button class="tr-opt" data-i="${i}"><span class="tr-radio"></span><span>${E(t)}</span></button>`).join("")}</div>
      <div class="tr-sure" hidden><span class="tr-muted">How sure are you?</span>${btn("Sure", "primary")}${btn("Guessing")}</div><div class="tr-after"></div>`;
    qMenu(host, q, lr.course);
    let chosen = -1, answered = false;
    const opts = [...host.querySelectorAll(".tr-opt")], sure = host.querySelector(".tr-sure");
    opts.forEach(o => o.onclick = () => {
      if (answered) return;
      chosen = +o.dataset.i;
      opts.forEach(x => x.classList.toggle("sel", x === o));
      if (scored) { sure.hidden = false; sure.querySelector(".primary").focus({ preventScroll: true }); }
      else answer(false);
    });
    const [sureBtn, guessBtn] = sure.querySelectorAll("button");
    sureBtn.onclick = () => answer(true); guessBtn.onclick = () => answer(false);
    const onKey = e => {
      if (answered || !document.body.contains(host)) { document.removeEventListener("keydown", onKey); return; }
      const n = parseInt(e.key, 10);
      if (n >= 1 && n <= opts.length) opts[n - 1].click();
    };
    document.addEventListener("keydown", onKey);
    function answer(isSure) {
      if (answered || chosen < 0) return;
      answered = true; sure.hidden = true;
      const right = chosen === q.answer;
      opts.forEach((o, k) => { o.disabled = true; o.classList.add(k === q.answer ? "right" : k === chosen ? "wrong" : "dim"); });
      const after = host.querySelector(".tr-after");
      let verdict = right ? "Right" : "Not quite";
      if (scored) {
        if (tally) { tally.asked++; if (right && isSure) tally.right++; }
        T.add({ Kind: "answer", Question: q.id, Concept: q.concept || "", Correct: right, Sure: isSure }, lr.course);
        verdict = right ? (isSure ? "Right, and you were sure" : "Right, but you were guessing") : isSure ? "Not quite, and you were sure" : "Not quite";
        const st = T.stateOf(lr.course, q.id, false);
        verdict += ` <span class="tr-muted">· back ${daysText(st.due)}</span>`;
      }
      after.innerHTML = `<p class="tr-verdict ${right ? "ok" : "bad"}">${verdict}</p>${q.why ? `<p class="tr-why">${E(q.why)}</p>` : ""}`;
      if ((q.refs || []).length) {
        const a = document.createElement("button"); a.className = "tr-link"; a.textContent = "Show in book";
        a.onclick = () => showInBook(lr, q.refs[0]); after.append(a);
      }
      nextButton(host, next);
    }
  }

  // ----- Put in order -----
  function orderStep(host, lr, q, label, next, tally) {
    const items = q.items || [];
    const shown = T.shuffle(items.map((t, i) => ({ t, i })));
    const picked = [];
    host.innerHTML = `<div class="tr-step">${E(label)}</div><h2 class="tr-q">${E(q.prompt || "Put these in order")}</h2><p class="tr-muted">Tap them in order, first to last. Tap a placed one to take it back.</p>
      <ol class="tr-order"></ol><div class="tr-pool"></div><div class="tr-after"></div>`;
    qMenu(host, q, lr.course);
    const list = host.querySelector(".tr-order"), pool = host.querySelector(".tr-pool");
    let checked = false;
    const paint = () => {
      list.innerHTML = picked.map((x, k) => `<li><button class="tr-chip on" data-k="${k}">${E(x.t)}</button></li>`).join("");
      pool.innerHTML = shown.filter(x => !picked.includes(x)).map(x => `<button class="tr-chip" data-i="${x.i}">${E(x.t)}</button>`).join("");
      if (!checked && picked.length === items.length && !host.querySelector(".tr-check")) {
        const c = document.createElement("div"); c.className = "tr-act tr-check"; c.innerHTML = btn("Check", "primary");
        c.querySelector("button").onclick = check; host.querySelector(".tr-after").before(c);
      }
      if (picked.length < items.length) host.querySelector(".tr-check")?.remove();
    };
    list.onclick = e => { const b = e.target.closest("[data-k]"); if (!b || checked) return; picked.splice(+b.dataset.k, 1); paint(); };
    pool.onclick = e => { const b = e.target.closest("[data-i]"); if (!b || checked) return; picked.push(shown.find(x => x.i === +b.dataset.i)); paint(); };
    paint();
    function check() {
      checked = true;
      host.querySelector(".tr-check").remove();
      const right = picked.every((x, k) => x.i === k);
      if (tally) { tally.asked++; if (right) tally.right++; }
      T.add({ Kind: "answer", Question: q.id, Concept: q.concept || "", Correct: right, Sure: true }, lr.course);
      list.querySelectorAll("button").forEach((b, k) => { b.disabled = true; b.classList.add(picked[k].i === k ? "right" : "wrong"); });
      host.querySelector(".tr-after").innerHTML = `<p class="tr-verdict ${right ? "ok" : "bad"}">${right ? "Right order" : "Not quite. The order is:"}</p>` +
        (right ? "" : `<ol class="tr-answer">${items.map(t => `<li>${E(t)}</li>`).join("")}</ol>`) + (q.why ? `<p class="tr-why">${E(q.why)}</p>` : "");
      nextButton(host, next);
    }
  }

  // ----- Memory cards -----
  function cardStep(host, lr, q, label, next, tally) {
    const front = q.front || q.prompt || "", backText = q.back || (q.options || [])[q.answer] || "";
    host.innerHTML = `<div class="tr-step">${E(label)}</div>
      <button class="tr-card"><span class="tr-card-f">${E(front)}</span><span class="tr-card-b" hidden>${E(backText)}</span><small class="tr-muted">Tap to flip</small></button>
      <p class="tr-muted tr-center">Grade it, flipped or not:</p>
      <div class="tr-grades"><button data-g="0" class="tr-btn g0">✕<small>Again</small></button><button data-g="1" class="tr-btn g1">✓<small>Know it</small></button><button data-g="2" class="tr-btn g2">✓✓<small>Know it well</small></button></div><div class="tr-after"></div>`;
    qMenu(host, q, lr.course);
    const card = host.querySelector(".tr-card");
    card.onclick = () => { host.querySelector(".tr-card-b").hidden = false; card.classList.add("flipped"); card.querySelector("small").hidden = true; };
    let graded = false;
    host.querySelectorAll("[data-g]").forEach(b => b.onclick = () => {
      if (graded) return;
      graded = true;
      const g = +b.dataset.g;
      if (tally) { tally.asked++; if (g > 0) tally.right++; }
      T.add({ Kind: "answer", Question: q.id, Concept: q.concept || "", Correct: g > 0, Sure: g > 0, Grade: g }, lr.course);
      card.click();
      host.querySelectorAll("[data-g]").forEach(x => { x.disabled = true; x.classList.toggle("sel", x === b); });
      const st = T.stateOf(lr.course, q.id, true);
      host.querySelector(".tr-after").innerHTML = `<p class="tr-muted">Back ${daysText(st.due)}.</p>${q.why ? `<p class="tr-why">${E(q.why)}</p>` : ""}`;
      setTimeout(next, g === 0 ? 1100 : 650);
    });
  }

  function questionStep(lr, q, label, key, tally) {
    return { key, run: (host, next) => q.kind === "card" ? cardStep(host, lr, q, label, next, tally) : q.kind === "order" ? orderStep(host, lr, q, label, next, tally) : mcStep(host, lr, q, label, true, next, tally) };
  }

  // ----- Main ideas (one summary per section) -----
  function mainIdeasStep(host, lr, next, tally) {
    const ideas = (lr.lesson.mainIdeas || []).filter(m => (m.right || "").trim());
    const sets = ideas.map(m => T.shuffle([{ t: m.right, ok: true }, ...(m.misses || []).map(t => ({ t, ok: false }))]));
    host.innerHTML = `<div class="tr-step">Main ideas</div><h2 class="tr-q">What is each section really saying?</h2><p class="tr-muted">Pick the summary that fits each section best. Two are only slightly off.</p>` +
      ideas.map((m, i) => `<div class="tr-sec" data-s="${i}"><div class="tr-sec-h"><span>${i + 1}</span>${E(m.heading || (lr.lesson.parts[m.part] || {}).title || "Section " + (i + 1))}</div>` +
        sets[i].map((o, k) => `<button class="tr-opt" data-k="${k}"><span class="tr-radio"></span><span>${E(o.t)}</span></button>`).join("") + `<div class="tr-sec-after"></div></div>`).join("") +
      `<div class="tr-act"><button class="tr-btn primary" disabled>Check</button></div><p class="tr-verdict"></p>`;
    const pick = {};
    const check = host.querySelector(".tr-act button");
    host.querySelectorAll(".tr-sec").forEach(sec => sec.onclick = e => {
      const o = e.target.closest(".tr-opt");
      if (!o || o.disabled) return;
      pick[sec.dataset.s] = +o.dataset.k;
      sec.querySelectorAll(".tr-opt").forEach(x => x.classList.toggle("sel", x === o));
      check.disabled = Object.keys(pick).length < ideas.length;
      if (check.disabled) sec.nextElementSibling?.scrollIntoView({ behavior: T.reader.isEink() ? "auto" : "smooth", block: "nearest" });
    });
    check.onclick = () => {
      let right = 0;
      host.querySelectorAll(".tr-sec").forEach(sec => {
        const i = +sec.dataset.s, k = pick[i], ok = sets[i][k].ok;
        if (ok) right++;
        sec.querySelectorAll(".tr-opt").forEach((x, j) => {
          x.disabled = true;
          if (sets[i][j].ok) x.classList.add("right"); else if (j === k) x.classList.add("wrong"); else x.hidden = true;
        });
        const m = ideas[i], after = sec.querySelector(".tr-sec-after");
        after.innerHTML = ok ? "" : `<p class="tr-why">${E(m.why || "")}</p>`;
        const find = (m.find || "").trim() || (lr.lesson.parts[m.part] || {}).start || "";
        const a = document.createElement("button"); a.className = "tr-link"; a.textContent = ok ? "✓ Right · Show in book" : "Show in book";
        a.onclick = () => showInBook(lr, { part: m.part, find }); after.append(a);
      });
      const score = Math.round((100 * right) / Math.max(1, ideas.length));
      if (tally) { tally.asked += ideas.length; tally.right += right; }
      T.add({ Kind: "mainideas", Lesson: lr.index, Grade: score, Correct: score >= 70, Note: lr.lesson.title || "" }, lr.course);
      host.querySelector(".tr-verdict").textContent = `${right} of ${ideas.length} section${ideas.length === 1 ? "" : "s"} right`;
      check.parentElement.remove();
      nextButton(host, next);
    };
  }

  // ----- Timeline in a lesson: type each year -----
  function timelineTypeStep(host, items, next, tally) {
    host.innerHTML = `<div class="tr-step">Timeline</div><h2 class="tr-q">When did each happen?</h2><p class="tr-muted">Type a year (like 4 BC, 30 or AD 30) and press Go.</p><div class="tr-tl-rows"></div><div class="tr-tl-line"></div>`;
    const rows = host.querySelector(".tr-tl-rows");
    const placed = {};
    items.forEach((x, i) => {
      const r = document.createElement("div"); r.className = "tr-tl-row";
      r.innerHTML = `<div class="tr-tl-text">${E(x.ev.text)}</div><form class="tr-tl-in"><input inputmode="text" autocomplete="off" enterkeyhint="go" aria-label="Year" placeholder="Year"><button class="tr-btn">Go</button></form><div class="tr-tl-res"></div>`;
      rows.append(r);
      r.querySelector("form").onsubmit = e => {
        e.preventDefault();
        if (i in placed) return;
        const inp = r.querySelector("input"), y = T.parseYear(inp.value);
        if (y == null) { inp.classList.add("bad"); return; }
        placed[i] = y;
        const off = Math.abs(y - x.ev.year), ok = off <= (x.ev.within ?? 10);
        if (tally) { tally.asked++; if (ok) tally.right++; }
        T.add({ Kind: "timeline", Question: x.ev.text, Correct: ok, Grade: off, Lesson: x.l.index }, x.book);
        inp.disabled = true; r.querySelector("button").remove();
        r.classList.add(ok ? "ok" : "bad");
        r.querySelector(".tr-tl-res").innerHTML = ok ? `✓ ${E(T.yearLabel(x.ev.year))}` : `✕ It was ${E(T.yearLabel(x.ev.year))}${x.ev.note ? ` · ${E(x.ev.note)}` : ""}`;
        const nextIn = rows.querySelectorAll("input:not(:disabled)")[0];
        if (nextIn) nextIn.focus(); else done();
      };
    });
    setTimeout(() => rows.querySelector("input")?.focus(), 50);
    function done() {
      const sorted = items.map((x, i) => ({ x, i })).sort((a, b) => a.x.ev.year - b.x.ev.year);
      host.querySelector(".tr-tl-line").innerHTML = `<h3 class="tr-h">On the timeline</h3><ol class="tr-vline">${sorted.map(({ x, i }) => `<li class="${Math.abs(placed[i] - x.ev.year) <= (x.ev.within ?? 10) ? "ok" : "bad"}"><b>${E(T.yearLabel(x.ev.year))}</b>${E(x.ev.text)}</li>`).join("")}</ol>`;
      nextButton(host, next);
    }
  }

  // ---------- A lesson ----------

  function runLesson(lr) {
    const id = lr.course, l = lr.lesson, done = T.stepsDone(id, lr.index), hidden = T.hidden(id);
    const tally = { asked: 0, right: 0 };
    const steps = [];
    const resuming = done.size > 0;
    steps.push({ key: null, run: (host, next) => {
      host.innerHTML = `<div class="tr-step">${E(T.course(id)?.Title || "")} · chapter ${lr.pack.chapter}, lesson ${lr.number} of ${lr.pack.lessons.length}</div><h2 class="tr-title">${E(l.title || "Lesson " + lr.number)}</h2>` +
        (resuming ? `<p>Welcome back. You'll pick up where you left off.</p>` : "") +
        (l.bigQuestions.length ? `<h3 class="tr-h">Big questions</h3><ul class="tr-list">${l.bigQuestions.map(b => `<li>${E(b)}</li>`).join("")}</ul>` : "") +
        (l.parts.length ? `<h3 class="tr-h">Sections</h3><ol class="tr-list">${l.parts.map(p => `<li>${E(p.title)}${(p.watch || []).length ? `<small>Watch for: ${E(p.watch.join(" · "))}</small>` : ""}</li>`).join("")}</ol>` : "");
      nextButton(host, next, resuming ? "Continue" : "Begin");
    } });
    if (!done.has("reading")) {
      const warm = l.warmup.filter(q => !hidden.has(q.id) && !done.has("warm:" + q.id));
      warm.forEach((q, n) => steps.push({ key: "warm:" + q.id, run: (host, next) => mcStep(host, lr, q, `Warm-up ${n + 1} of ${warm.length} · just a guess`, false, next) }));
      steps.push({ key: "reading", run: (host, next) => readingStep(host, lr, next) });
    }
    if (!done.has("mainideas") && (l.mainIdeas || []).some(m => (m.right || "").trim())) steps.push({ key: "mainideas", run: (host, next) => mainIdeasStep(host, lr, next, tally) });
    const practice = l.practice.filter(q => !hidden.has(q.id) && !done.has("practice:" + q.id));
    practice.forEach((q, n) => steps.push(questionStep(lr, q, `Practice ${n + 1} of ${practice.length}`, "practice:" + q.id, tally)));
    const cards = l.cards.filter(q => !hidden.has(q.id) && !done.has("card:" + q.id));
    cards.forEach((q, n) => steps.push({ key: "card:" + q.id, run: (host, next) => cardStep(host, lr, q, `Memory card ${n + 1} of ${cards.length}`, next, tally) }));
    if (l.endsChapter || l.quiz.length) {
      const quiz = l.quiz.filter(q => !hidden.has(q.id) && !done.has("quiz:" + q.id));
      quiz.forEach((q, n) => steps.push(questionStep(lr, q, `Chapter quiz ${n + 1} of ${quiz.length}`, "quiz:" + q.id, tally)));
    }
    if (!done.has("timeline")) {
      const tl = T.timelineFor(lr);
      if (tl.length >= 2) steps.push({ key: "timeline", run: (host, next) => timelineTypeStep(host, tl, next, tally) });
    }
    const review = T.dueReview(id, T.config.ReviewCap || 8).filter(x => !done.has("review:" + x.q.id)).sort((a, b) => (a.q.kind === "card" ? 0 : 1) - (b.q.kind === "card" ? 0 : 1));
    review.forEach((x, n) => steps.push(questionStep(x.l, x.q, `Review ${n + 1} of ${review.length}`, "review:" + x.q.id, tally)));
    steps.push({ key: null, run: host => {
      T.add({ Kind: "lessonDone", Lesson: lr.index }, id);
      const nx = T.nextLesson(id);
      host.innerHTML = `<h2 class="tr-title">Lesson finished</h2><p class="tr-big">${tally.asked ? `${tally.right} of ${tally.asked} right and sure (this sitting).` : "Nicely done."}</p>
        <p>Questions you missed or guessed come back tomorrow; ones you knew come back after a longer gap. It counts on your computers too.</p>${nx ? `<p class="tr-muted">Next: chapter ${nx.pack.chapter} · ${E(nx.lesson.title || "Lesson " + nx.number)}</p>` : ""}`;
      nextButton(host, back, "Close");
    } });
    player(l.title || "Lesson", steps, { onStep: key => { if (!T.stepsDone(id, lr.index).has(key)) T.add({ Kind: "lessonStep", Lesson: lr.index, Question: key }, id); } });
  }
  function readingStep(host, lr, next) {
    const [from, to] = T.lessonRange(lr);
    const l = lr.lesson;
    host.innerHTML = `<div class="tr-step">Reading</div><h2 class="tr-q">Read: ${E(l.title || "this lesson")}</h2>` +
      (l.parts.length ? `<ol class="tr-list">${l.parts.map(p => `<li>${E(p.title)}${(p.watch || []).length ? `<small>Watch for: ${E(p.watch.join(" · "))}</small>` : ""}</li>`).join("")}</ol>` : "") +
      (from < 0 ? `<p class="tr-bad">This lesson's reading wasn't found in the book on this phone. Sync again, or read it on your computer.</p>` : `<p class="tr-muted">About ${Math.max(1, Math.round(T.B(lr.course).book.slice(from, to).reduce((n, p) => n + T.paraText(p).split(/\s+/).length, 0) / 230))} minutes. Tap the right side or swipe to turn pages.</p>`) +
      `<div class="tr-act">${from >= 0 ? btn("Open the reading", "primary") : ""}${btn("I've read it", from >= 0 ? "" : "primary")}</div>`;
    const [openB, doneB] = from >= 0 ? host.querySelectorAll(".tr-act button") : [null, host.querySelector(".tr-act button")];
    const open = () => {
      T.openReader({ courseId: lr.course, from, to, title: l.title || `Chapter ${lr.pack.chapter}`, chapter: lr.pack.chapter, doneLabel: "Done reading →", onDone: next, requestClose: back });
      openLayer(T.closeReader);
    };
    if (openB) { openB.onclick = open; setTimeout(open, 150); }
    doneB.onclick = next;
  }

  // ---------- Practice ----------

  function finishScreen(title, tally, extra = "") {
    return { key: null, run: host => {
      host.innerHTML = `<h2 class="tr-title">${E(title)}</h2><p class="tr-big">${tally.asked ? `${tally.right} of ${tally.asked} right.` : "Nothing answered this time."}</p>${extra}`;
      nextButton(host, back, "Close");
    } };
  }
  function runBooksPractice() {
    const tally = { asked: 0, right: 0 };
    const list = ids().flatMap(id => T.dueReview(id).filter(x => x.q.kind !== "card")).slice(0, UI.books);
    if (!list.length) { openReviewChooser(); return; }
    const steps = list.map((x, n) => questionStep(x.l, x.q, `Question ${n + 1} of ${list.length} · ${T.course(x.l.course)?.Title || ""}, ch. ${x.l.pack.chapter}`, null, tally));
    steps.push(finishScreen("Books practice done", tally, `<p>Missed ones come back tomorrow.</p>`));
    player("Books practice", steps);
  }
  /** Nothing due: review a lesson you choose. */
  function openReviewChooser() {
    sheet("Nothing due right now", `<p class="tr-muted">Review a lesson you've done instead:</p>` + courses().map(c => {
      const done = T.doneLessons(c.Id), ls = T.B(c.Id).lessons.filter(l => done.has(l.index));
      return ls.length ? `<h3 class="tr-h">${E(c.Title)}</h3>` + ls.slice(-12).reverse().map(l => `<button class="tr-rowbtn" data-review="${E(c.Id)}|${l.index}">Ch. ${l.pack.chapter} · ${E(l.lesson.title || "Lesson " + l.number)}</button>`).join("") : "";
    }).join(""), el => el.onclick = e => {
      const b = e.target.closest("[data-review]");
      if (!b) return;
      const [cid, idx] = b.dataset.review.split("|");
      back();
      setTimeout(() => reviewLesson(T.B(cid).lessons[+idx]), 60);
    });
  }
  function reviewLesson(lr) {
    const tally = { asked: 0, right: 0 }, hidden = T.hidden(lr.course);
    const qs = [...lr.lesson.practice, ...lr.lesson.quiz].filter(q => q.kind !== "card" && !hidden.has(q.id));
    if (!qs.length) { note("This lesson has no questions to review."); return; }
    const steps = qs.map((q, n) => questionStep(lr, q, `Question ${n + 1} of ${qs.length}`, null, tally));
    steps.push(finishScreen("Review done", tally));
    player(lr.lesson.title || "Review", steps);
  }
  function runFlashcards() {
    const tally = { asked: 0, right: 0 };
    let list = ids().flatMap(id => T.dueReview(id).filter(x => x.q.kind === "card"));
    let title = "Flashcards";
    if (!list.length) { list = ids().flatMap(id => T.hardestCardsToday(id)); title = "Today's hardest cards"; }
    list = list.slice(0, UI.cards);
    if (!list.length) { note("No cards due. Finish a lesson and its cards come up here."); return; }
    const steps = list.map((x, n) => ({ key: null, run: (host, next) => cardStep(host, x.l, x.q, `Card ${n + 1} of ${list.length}`, next, tally) }));
    steps.push({ key: null, run: host => {
      T.add({ Kind: "cardReviewDone", Lesson: -1 }, list[0].l.course);
      host.innerHTML = `<h2 class="tr-title">Flashcards finished</h2><p class="tr-big">${tally.asked ? `${tally.right} of ${tally.asked} known.` : "Nothing graded this time."}</p><p>Cards you marked ✕ come back tomorrow; ✓ after a while; ✓✓ after a long while.</p>`;
      nextButton(host, back, "Close");
    } });
    player(title, steps);
  }
  function runQuestions(title, list, finish) {
    const tally = { asked: 0, right: 0 };
    const steps = [{ key: null, run: (host, next) => { host.innerHTML = `<h2 class="tr-title">${E(title)}</h2><p>${plural(list.length, "question")}. Each names its book and chapter.</p>`; nextButton(host, next, "Begin"); } }];
    list.forEach((x, n) => steps.push(questionStep(x.l, x.q, `Question ${n + 1} of ${list.length} · ${T.course(x.l.course)?.Title || ""}, ch. ${x.l.pack.chapter}`, null, tally)));
    steps.push({ key: null, run: host => { finish(host, tally); nextButton(host, back, "Close"); } });
    player(title, steps);
  }
  function runExam(id) {
    const exam = T.examDue(id);
    if (!exam) return;
    runQuestions(exam.title, exam.questions, (host, t) => {
      const passed = t.asked > 0 && t.right * 10 >= t.asked * 7;
      T.add({ Kind: "examDone", Note: exam.key, Correct: passed, Lesson: -1 }, id);
      T.add({ Kind: "lessonDone", Lesson: -1, Note: "exam:" + exam.key }, id);
      host.innerHTML = `<h2 class="tr-title">${passed ? "Passed" : "Not passed yet"}</h2><p class="tr-big">${t.right} of ${t.asked} right and sure.</p><p>${passed ? (exam.key === "final" ? "You've finished the book. Its questions now join All books." : "On to the next part of the book.") : "The questions you missed come back in review. The exam comes back as tomorrow's lesson."}</p>`;
    });
  }
  function runAllBooks() {
    const list = T.allBooksRound(20);
    if (!list.length) { note("Finish a book (its final exam) and its questions come here."); return; }
    runQuestions("All books", list, (host, t) => {
      T.add({ Kind: "allBooksDone" }, list[0].l.course);
      host.innerHTML = `<h2 class="tr-title">Round finished</h2><p class="tr-big">${t.right} of ${t.asked} right and sure.</p>`;
    });
  }

  // ----- Matching (timeline and Greek): tap one on the left, then its match on the right -----

  function matchBoard(host, { left, right, title, hint, onPair, onCheck, checkMode }) {
    host.innerHTML = `<div class="tr-step">${E(title)}</div><p class="tr-muted">${E(hint)}</p><div class="tr-match"><div class="tr-ml"></div><div class="tr-mr"></div></div><div class="tr-after"></div>`;
    const L = host.querySelector(".tr-ml"), Rt = host.querySelector(".tr-mr");
    L.innerHTML = left.map((x, i) => `<button class="tr-mi" data-l="${i}">${x.html}</button>`).join("");
    Rt.innerHTML = right.map((x, i) => `<button class="tr-mi" data-r="${i}">${x.html}</button>`).join("");
    const state = { sel: null, pairs: {} }; // left index → right index
    const paint = () => {
      L.querySelectorAll("[data-l]").forEach(b => { const i = +b.dataset.l; b.classList.toggle("sel", state.sel === i); b.dataset.pair = i in state.pairs ? Object.keys(state.pairs).indexOf(String(i)) + 1 : ""; });
      const used = Object.fromEntries(Object.entries(state.pairs).map(([l, r]) => [r, l]));
      Rt.querySelectorAll("[data-r]").forEach(b => { const r = +b.dataset.r; b.dataset.pair = r in used ? Object.keys(state.pairs).indexOf(used[r]) + 1 : ""; });
    };
    host.querySelector(".tr-match").onclick = e => {
      const b = e.target.closest(".tr-mi");
      if (!b || b.disabled) return;
      if (b.dataset.l != null) {
        const i = +b.dataset.l;
        if (checkMode) { if (i in state.pairs) { delete state.pairs[i]; onCheck(state); } state.sel = i; }
        else state.sel = state.sel === i ? null : i;
      } else if (state.sel != null) {
        const r = +b.dataset.r;
        if (checkMode) {
          for (const k of Object.keys(state.pairs)) if (state.pairs[k] === r) delete state.pairs[k];
          state.pairs[state.sel] = r;
          const nextFree = left.findIndex((_, k) => !(k in state.pairs) && !L.querySelector(`[data-l="${k}"]`).disabled);
          state.sel = nextFree >= 0 ? nextFree : null;
          onCheck(state);
        } else onPair(state.sel, r, b, L.querySelector(`[data-l="${state.sel}"]`), state);
      }
      paint();
    };
    paint();
    return { state, paint, L, R: Rt };
  }

  function runTimelinePractice() {
    const books = ids();
    if (!T.learnedEvents(books).length) { note("Timeline events show up here after the lessons that teach them."); return; }
    let round = 0;
    const tally = { asked: 0, right: 0 };
    const roundStep = () => ({ key: null, run: (host, next) => {
      round++;
      const items = T.timelineRustiest(books, Math.max(2, UI.timeline));
      // The real years plus two that belong to none of them.
      const years = items.map(x => x.ev.year);
      const lo = Math.min(...years), hi = Math.max(...years), span = Math.max(20, hi - lo);
      const extras = [];
      for (let k = 0; extras.length < 2 && k < 200; k++) {
        const y = Math.round(lo - span * 0.3 + Math.random() * span * 1.6);
        if (y === 0 || extras.includes(y) || items.some(x => Math.abs(y - x.ev.year) <= Math.max(x.ev.within ?? 10, 3))) continue;
        extras.push(y);
      }
      const dates = [...new Set([...years, ...extras])].sort((a, b) => a - b);
      const recorded = new Set();
      const board = matchBoard(host, {
        title: `Timeline · round ${round}`, hint: "Tap an event, then its year. Check shows which are wrong (not the answers); fix them and check again.", checkMode: true,
        left: items.map(x => ({ html: E(x.ev.text) })), right: dates.map(y => ({ html: `<b>${E(T.yearLabel(y))}</b>` })),
        onCheck: st => { checkB.disabled = Object.keys(st.pairs).length < items.length; }
      });
      const act = document.createElement("div"); act.className = "tr-act"; act.innerHTML = btn("Check", "primary") + btn("Done");
      host.querySelector(".tr-after").before(act);
      const [checkB, doneB] = act.querySelectorAll("button");
      checkB.disabled = true;
      doneB.onclick = () => { T.add({ Kind: "timelineDone", Lesson: -1 }, books[0]); back(); };
      checkB.onclick = () => {
        let wrong = 0;
        items.forEach((x, i) => {
          const y = dates[board.state.pairs[i]], off = Math.abs(y - x.ev.year), ok = off <= (x.ev.within ?? 10);
          const lb = board.L.querySelector(`[data-l="${i}"]`);
          if (!recorded.has(i)) { recorded.add(i); tally.asked++; if (ok) tally.right++; T.add({ Kind: "timeline", Question: x.ev.text, Correct: ok, Grade: off, Lesson: x.l.index }, x.book); }
          if (ok) { lb.disabled = true; lb.classList.add("right"); board.R.querySelector(`[data-r="${board.state.pairs[i]}"]`).classList.add("right"); board.R.querySelector(`[data-r="${board.state.pairs[i]}"]`).disabled = true; }
          else { wrong++; lb.classList.add("wrong"); setTimeout(() => lb.classList.remove("wrong"), 1400); delete board.state.pairs[i]; }
        });
        board.paint();
        checkB.disabled = true;
        const after = host.querySelector(".tr-after");
        if (wrong) after.innerHTML = `<p class="tr-verdict bad">${wrong} still wrong. Try again.</p>`;
        else {
          after.innerHTML = `<p class="tr-verdict ok">All placed.</p><ol class="tr-vline">${items.slice().sort((a, b) => a.ev.year - b.ev.year).map(x => `<li><b>${E(T.yearLabel(x.ev.year))}</b>${E(x.ev.text)}</li>`).join("")}</ol>`;
          act.remove();
          const row = document.createElement("div"); row.className = "tr-act"; row.innerHTML = btn("Another round →", "primary") + btn("Done");
          after.append(row);
          const [more, fin] = row.querySelectorAll("button");
          more.onclick = () => { steps.push(roundStep()); next(); };
          fin.onclick = () => { T.add({ Kind: "timelineDone", Lesson: -1 }, books[0]); back(); };
        }
      };
    } });
    const steps = [roundStep()];
    player("Timeline", steps);
  }

  function runGreek(kind) {
    const name = { vocab: "Greek words", alphabet: "Alphabet", concepts: "Greek concepts" }[kind];
    let round = 0, right = 0, asked = 0;
    const roundStep = () => ({ key: null, run: (host, next) => {
      round++;
      const cards = T.greekRound(kind, Math.max(3, UI.greek));
      if (!cards.length) { host.innerHTML = `<p>Nothing to practise yet.</p>`; nextButton(host, back, "Close"); return; }
      const extras = T.greekExtras(kind, cards, 2);
      const answers = T.shuffle([...cards, ...extras]);
      const missed = new Set(), done = new Set();
      matchBoard(host, {
        title: `${name} · round ${round}`, hint: "Tap a Greek word, then its meaning.",
        left: cards.map(c => ({ html: `<span class="tr-greek">${E(c.Greek)}</span>` })), right: answers.map(c => ({ html: E(c.English) })),
        onPair: (l, r, rb, lb, st) => {
          const c = cards[l], ok = answers[r] === c || T.plain(answers[r].English).trim() === T.plain(c.English).trim();
          if (ok) {
            if (!missed.has(c)) { T.greekRecord(kind, c, true); right++; }
            asked++;
            lb.disabled = rb.disabled = true; lb.classList.add("right"); rb.classList.add("right");
            done.add(c); st.sel = null;
            if (done.size === cards.length) finish();
          } else {
            if (!missed.has(c)) { missed.add(c); T.greekRecord(kind, c, false); }
            rb.classList.add("wrong"); lb.classList.add("wrong");
            setTimeout(() => { rb.classList.remove("wrong"); lb.classList.remove("wrong"); }, 700);
          }
        }
      });
      const act = document.createElement("div"); act.className = "tr-act"; act.innerHTML = btn("Done");
      host.append(act);
      act.querySelector("button").onclick = back;
      function finish() {
        act.innerHTML = btn("Another round →", "primary") + btn("Done");
        const [more, fin] = act.querySelectorAll("button");
        host.querySelector(".tr-after").innerHTML = `<p class="tr-verdict ${missed.size ? "bad" : "ok"}">${cards.length - missed.size} of ${cards.length} right first time.</p>` +
          (missed.size ? `<p class="tr-muted">Missed ones come back in 30 minutes.</p>` : "");
        more.onclick = () => { steps.push(roundStep()); next(); };
        fin.onclick = back;
        more.focus();
      }
    } });
    const steps = [roundStep()];
    player(name, steps, { onClose: () => note(asked ? `${right} of ${asked} right first time.` : "Saved.") });
  }

  // ----- Greek concept quiz (from your Greek lessons) -----
  function openQuizList() {
    sheet("Concept quiz", `<button class="tr-rowbtn strong" data-q="review">Mixed review · 10 questions from every lesson</button>` +
      T.greek.quizzes.map((p, i) => `<button class="tr-rowbtn" data-q="${i}">Lesson ${E(p.Lesson)}${p.Title ? " · " + E(p.Title) : ""}<small>${plural(p.Questions.length, "question")}</small></button>`).join(""),
      el => el.onclick = e => {
        const b = e.target.closest("[data-q]");
        if (!b) return;
        back();
        const qs = b.dataset.q === "review" ? T.shuffle(T.greek.quizzes.flatMap(p => p.Questions.map(q => ({ p, q })))).slice(0, 10) : T.greek.quizzes[+b.dataset.q].Questions.map(q => ({ p: T.greek.quizzes[+b.dataset.q], q }));
        setTimeout(() => runQuiz(b.dataset.q === "review" ? "Mixed review" : "Lesson " + qs[0].p.Lesson, qs), 60);
      });
  }
  function runQuiz(title, qs) {
    let right = 0;
    const steps = qs.map(({ p, q }, n) => ({ key: null, run: (host, next) => {
      const multi = q.Correct.length > 1;
      host.innerHTML = `<div class="tr-step">Question ${n + 1} of ${qs.length}${multi ? " · pick every right answer" : ""}</div><h2 class="tr-q">${E(q.Prompt)}</h2>
        <div class="tr-opts">${q.Options.map((t, i) => `<button class="tr-opt${multi ? " multi" : ""}" data-i="${i}"><span class="tr-radio"></span><span>${E(t)}</span></button>`).join("")}</div>
        <div class="tr-act"><button class="tr-btn primary" disabled>Check</button></div><div class="tr-after"></div>`;
      const picked = new Set(), opts = [...host.querySelectorAll(".tr-opt")], check = host.querySelector(".tr-act button");
      opts.forEach(o => o.onclick = () => {
        const i = +o.dataset.i;
        if (multi) { picked.has(i) ? picked.delete(i) : picked.add(i); } else { picked.clear(); picked.add(i); }
        opts.forEach(x => x.classList.toggle("sel", picked.has(+x.dataset.i)));
        check.disabled = !picked.size;
      });
      check.onclick = () => {
        const want = new Set(q.Correct), ok = want.size === picked.size && [...want].every(i => picked.has(i));
        if (ok) right++;
        opts.forEach((o, i) => { o.disabled = true; o.classList.add(...(want.has(i) ? (picked.has(i) ? ["right"] : ["right", "missed"]) : picked.has(i) ? ["wrong"] : ["dim"])); });
        check.parentElement.remove();
        const sec = p.Sections.find(s => s.Id === q.Section);
        host.querySelector(".tr-after").innerHTML = `<p class="tr-verdict ${ok ? "ok" : "bad"}">${ok ? "Right" : "Not quite"}</p>${q.Explanation ? `<p class="tr-why">${E(q.Explanation)}</p>` : ""}` +
          (sec ? `<details class="tr-secread"><summary>Read: ${E(sec.Heading)}</summary>${sec.Paragraphs.map(x => `<p>${E(x)}</p>`).join("")}</details>` : "");
        nextButton(host, next);
      };
    } }));
    steps.push({ key: null, run: host => { host.innerHTML = `<h2 class="tr-title">Quiz finished</h2><p class="tr-big">${right} of ${qs.length} right.</p>`; nextButton(host, back, "Close"); } });
    player(title, steps);
  }

  // ---------- A book: lessons, reading and highlights ----------

  function openBook(id) {
    const c = T.course(id), b = T.B(id), done = T.doneLessons(id);
    const next = T.nextLesson(id), pct = Math.round(T.bookFraction(id) * 100);
    const hl = T.highlights(id);
    const chapters = b.packs.map(p => {
      const ls = b.lessons.filter(l => l.pack === p);
      return `<details class="tr-ch"${next && next.pack === p ? " open" : ""}><summary><b>Chapter ${p.chapter}</b> ${E(p.title || "")}<small>${ls.filter(l => done.has(l.index)).length} of ${ls.length} done</small></summary>` +
        `<button class="tr-rowbtn" data-read="ch|${p.chapter}">Read the chapter</button>` +
        ls.map(l => `<div class="tr-lrow${done.has(l.index) ? " done" : ""}"><span>${done.has(l.index) ? "✓" : l === next ? "→" : "·"} ${E(l.lesson.title || "Lesson " + l.number)}</span>` +
          (done.has(l.index) ? `<button class="tr-mini" data-rev="${l.index}">Review</button><button class="tr-mini" data-read="l|${l.index}">Read</button>` : l === next ? `<button class="tr-mini primary" data-go="${l.index}">${T.inProgress(l) ? "Continue" : "Start"}</button>` : "") + `</div>`).join("") + `</details>`;
    }).join("");
    sheet(c.Title, `<div class="tr-bookhead"><span class="tr-bar big"><span style="width:${pct}%"></span></span><small>${pct}% through the book${T.finished(id) ? " · finished" : ""}</small></div>` +
      (next ? `<button class="tr-btn primary wide" data-go="${next.index}">${T.inProgress(next) ? "Continue" : "Extra lesson"}: ${E(next.lesson.title || "Lesson " + next.number)}</button>` : "") +
      `<h3 class="tr-h">Chapters</h3>${chapters || `<p class="tr-muted">No chapter packs yet.</p>`}` +
      `<h3 class="tr-h">Highlights (${hl.length})</h3>` + (hl.length ? hl.slice(-30).reverse().map(h => `<button class="tr-hl ${{ yellow: "hl-y", green: "hl-g", blue: "hl-b" }[h.color] || "hl-y"}" data-hl="${h.chapter}|${h.para}"><span>“${E(h.quote)}”</span>${h.note ? `<small>${E(h.note)}</small>` : ""}<small>Chapter ${h.chapter}</small></button>`).join("") : `<p class="tr-muted">Select text while reading to highlight it.</p>`),
      el => el.onclick = e => {
        const t = e.target.closest("[data-go],[data-rev],[data-read],[data-hl]");
        if (!t) return;
        const lessonAt = i => b.lessons[+i];
        if (t.dataset.go != null) { back(); setTimeout(() => runLesson(lessonAt(t.dataset.go)), 60); }
        else if (t.dataset.rev != null) { back(); setTimeout(() => reviewLesson(lessonAt(t.dataset.rev)), 60); }
        else if (t.dataset.read) {
          const [kind, v] = t.dataset.read.split("|");
          const lr = kind === "l" ? lessonAt(v) : null, pack = lr ? lr.pack : b.packOf(+v);
          const [from, to] = lr ? T.lessonRange(lr) : T.chapterRange(id, pack);
          if (from < 0) { note("Couldn't find this in the book on this phone."); return; }
          T.openReader({ courseId: id, from, to, title: lr ? lr.lesson.title : `Chapter ${pack.chapter}`, chapter: pack.chapter, requestClose: back });
          openLayer(T.closeReader);
        } else if (t.dataset.hl) {
          const [ch, para] = t.dataset.hl.split("|").map(Number), pack = b.packOf(ch);
          if (!pack) return;
          const [from, to] = T.chapterRange(id, pack);
          T.openReader({ courseId: id, from, to, title: `Chapter ${ch}`, chapter: ch, goTo: para, requestClose: back });
          openLayer(T.closeReader);
        }
      });
  }

  // ---------- Sheets (settings, choosers) ----------

  function sheet(title, html, bind) {
    const el = document.createElement("div");
    el.className = "trs";
    el.innerHTML = `<div class="trs-card" role="dialog" aria-label="${E(title)}"><div class="trs-head"><b>${E(title)}</b><button class="tr-ic" data-x aria-label="Close">✕</button></div><div class="trs-body">${html}</div></div>`;
    document.body.append(el);
    let closed = false;
    openLayer(() => { if (!closed) { closed = true; el.remove(); } });
    el.querySelector("[data-x]").onclick = back;
    el.addEventListener("click", e => { if (e.target === el) back(); });
    if (bind) bind(el.querySelector(".trs-body"));
    return el;
  }
  function openSettings() {
    const P = T.reader.prefs;
    const num = (k, label, list) => `<div class="trr-row"><span>${label}</span><div class="tr-seg">${list.map(v => `<button data-ui="${k}" data-v="${v}" aria-pressed="${UI[k] === v}">${v}</button>`).join("")}</div></div>`;
    sheet("Training settings", `
      <div class="trr-row"><span>E-ink screen</span><div class="tr-seg">${["auto", "on", "off"].map(v => `<button data-eink="${v}" aria-pressed="${P.eink === v}">${v[0].toUpperCase() + v.slice(1)}</button>`).join("")}</div></div>
      <p class="tr-muted">For Boox and other e-ink readers: no animations, plain black on white, instant page turns. Auto turns it on when this looks like an e-ink device.</p>
      ${num("books", "Books practice", [5, 10, 15, 20])}${num("cards", "Flashcards", [10, 20, 30])}${num("timeline", "Timeline events", [4, 6, 8])}${num("greek", "Greek per round", [4, 6, 8])}
      <p class="tr-muted">Reading: tap Aa while reading for font, size, page color and margins.</p>
      <div class="tr-act">${btn("Sync now", "primary")}</div>
      <p class="tr-muted">This phone saves its progress as <b>${E(T.device)}</b> (its own file in Focus/training, the way each computer has one). Lesson days, days off and books are set on your computer.</p>
      <p class="tr-muted">${E(T.status.text || "")}</p>`, el => el.onclick = e => {
      const b = e.target.closest("button");
      if (!b) return;
      if (b.dataset.eink) { P.eink = b.dataset.eink; T.reader.savePrefs(); document.documentElement.classList.toggle("eink", T.reader.isEink()); }
      else if (b.dataset.ui) { UI[b.dataset.ui] = +b.dataset.v; saveUI(); }
      else if (b.classList.contains("primary")) { T.sync(); note("Syncing…"); }
      el.querySelectorAll("[data-eink]").forEach(x => x.setAttribute("aria-pressed", String(x.dataset.eink === P.eink)));
      el.querySelectorAll("[data-ui]").forEach(x => x.setAttribute("aria-pressed", String(UI[x.dataset.ui] === +x.dataset.v)));
    });
  }

  // ---------- The Training button in the header (shown once you're signed in) ----------

  function addButton() {
    const menu = $("menuBtn");
    if (!menu || $("trainBtn")) return;
    const b = document.createElement("button");
    b.className = "icon hidden"; b.id = "trainBtn"; b.setAttribute("aria-label", "Training");
    b.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 5.5C4 4.7 4.7 4 5.5 4H11v15H5.5c-.8 0-1.5.7-1.5 1.5z"/><path d="M20 5.5c0-.8-.7-1.5-1.5-1.5H13v15h5.5c.8 0 1.5.7 1.5 1.5z"/></svg>`;
    menu.before(b);
    b.onclick = openTraining;
    const mirror = () => {
      const on = !menu.classList.contains("hidden");
      b.classList.toggle("hidden", !on);
      if (on && !started && localStorage.getItem("focus-train-used")) { ensureRoot(); started = true; T.start(); }
    };
    new MutationObserver(mirror).observe(menu, { attributes: true, attributeFilter: ["class"] });
    mirror();
    if (location.hash === "#training") setTimeout(() => { if (!b.classList.contains("hidden")) openTraining(); }, 400);
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", addButton); else addButton();
})();
