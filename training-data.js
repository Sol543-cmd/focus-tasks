// Focus Training on your phone: the data. Reads the same training files Focus uses on your computers
// (Focus/training in OneDrive), keeps a copy on this device so you can study offline, and writes only
// this device's own progress file (training/progress-web-....json), the same way each computer has its
// own. So nothing you do here can overwrite a computer's work, and a lesson finished here counts there.
"use strict";

const T = window.T = {};

// ---------- Small helpers ----------

T.pad = n => String(n).padStart(2, "0");
T.ymd = d => `${d.getFullYear()}-${T.pad(d.getMonth() + 1)}-${T.pad(d.getDate())}`;
T.todayYmd = () => T.ymd(new Date());
T.startOfDay = d => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
T.addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
T.parseDay = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ""); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
/** A time in the form your computers write for local times (with the offset), so they read it as local. */
T.localIso = d => {
  const off = -d.getTimezoneOffset(), sign = off >= 0 ? "+" : "-", a = Math.abs(off);
  return `${T.ymd(d)}T${T.pad(d.getHours())}:${T.pad(d.getMinutes())}:${T.pad(d.getSeconds())}.${String(d.getMilliseconds()).padStart(3, "0")}${sign}${T.pad(Math.floor(a / 60))}:${T.pad(a % 60)}`;
};
/** Dates as your computers wrote them: "...Z" (UTC), with an offset (local) or with neither (local). */
T.date = s => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; };
T.shuffle = list => { const a = list.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
T.esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const DAY = 86400000;
/** A course's folder name, as your computers make it (letters, digits and hyphens; anything else becomes _). */
T.safe = s => Array.from(String(s)).map(c => (/[\p{L}\p{N}-]/u.test(c) ? c : "_")).join("");

// ---------- This device ----------

T.device = (() => {
  let id = localStorage.getItem("focus-train-device");
  if (!id) { id = "web-" + Math.random().toString(36).slice(2, 8); localStorage.setItem("focus-train-device", id); }
  return id;
})();
const MY_LOG = `progress-${T.device}.json`;

// ---------- The copy kept on this device (IndexedDB) ----------

const DB = (() => {
  let dbp = null;
  const open = () => dbp || (dbp = new Promise((ok, fail) => {
    const r = indexedDB.open("focus-training", 1);
    r.onupgradeneeded = () => { r.result.createObjectStore("files"); r.result.createObjectStore("kv"); };
    r.onsuccess = () => ok(r.result); r.onerror = () => fail(r.error);
  }));
  const tx = async (store, mode, fn) => {
    const db = await open();
    return new Promise((ok, fail) => { const t = db.transaction(store, mode); const req = fn(t.objectStore(store)); t.oncomplete = () => ok(req && req.result); t.onerror = () => fail(t.error); });
  };
  return {
    get: (store, k) => tx(store, "readonly", s => s.get(k)),
    put: (store, k, v) => tx(store, "readwrite", s => s.put(v, k)),
    del: (store, k) => tx(store, "readwrite", s => s.delete(k)),
    keys: store => tx(store, "readonly", s => s.getAllKeys()),
  };
})();
T.DB = DB;

// ---------- OneDrive (only Focus's training and Greek files) ----------

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0/me/drive/root:/";
/** The paths this page may read; it may only write its own progress file and the Greek progress files. */
const READ_OK = p => /^training\/(course\.json|progress-[A-Za-z0-9_-]{1,80}\.json|covers\/[A-Za-z0-9 ._-]{1,120}|[A-Za-z0-9_-]{1,80}\/(book\.json|ch\d{2,3}\.focuspack|media\/[A-Za-z0-9 ._-]{1,120}))$/.test(p)
  || /^(vocab|alphabet|concepts)\.json$/.test(p) || /^greek-concepts\/[A-Za-z0-9 ._-]{1,120}\.quizpack$/.test(p);
const WRITE_OK = p => p === "training/" + MY_LOG || /^(vocab|alphabet|concepts)\.json$/.test(p);
const LIST_OK = p => /^(training|training\/covers|training\/[A-Za-z0-9_-]{1,80}|training\/[A-Za-z0-9_-]{1,80}\/media|greek-concepts)$/.test(p);
function folderName() {
  const f = String((window.FOCUS_CONFIG || {}).folder || "Focus");
  if (!/^[A-Za-z0-9 _-]{1,64}$/.test(f)) throw new Error("Blocked: the folder name in config.js isn't allowed.");
  return f;
}
const enc = p => encodeURIComponent(folderName()) + "/" + p.split("/").map(encodeURIComponent).join("/");

T.remote = window.TRAINING_REMOTE || {
  async list(path) {
    if (!LIST_OK(path)) throw new Error("Blocked: " + path);
    let url = GRAPH_ROOT + enc(path) + ":/children?$select=name,lastModifiedDateTime,eTag,size,file,folder&$top=999";
    const out = [];
    while (url) {
      const res = await fetch(url, { headers: { Authorization: "Bearer " + await token() }, cache: "no-store" });
      if (res.status === 404) return [];
      if (!res.ok) throw new Error(`Couldn't list ${path} (${res.status})`);
      const j = await res.json();
      for (const it of j.value || []) out.push({ name: it.name, etag: it.eTag || it.lastModifiedDateTime, folder: !!it.folder, size: it.size || 0 });
      url = j["@odata.nextLink"] || null;
    }
    return out;
  },
  async get(path, binary) {
    if (!READ_OK(path)) throw new Error("Blocked: " + path);
    const res = await fetch(GRAPH_ROOT + enc(path) + ":/content", { headers: { Authorization: "Bearer " + await token() }, cache: "no-store" });
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Couldn't read ${path} (${res.status})`);
    return binary ? await res.blob() : await res.text();
  },
  async put(path, text) {
    if (!WRITE_OK(path)) throw new Error("Blocked: " + path);
    const res = await fetch(GRAPH_ROOT + enc(path) + ":/content", {
      method: "PUT", headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json" }, body: text
    });
    if (!res.ok) throw new Error(`Couldn't save ${path} (${res.status})`);
  }
};

// ---------- Everything Focus knows ----------

T.config = { Courses: [], WeekPlan: {}, DaysOff: {} };
T.events = [];        // every device's progress, merged
T.mine = [];          // this device's own events (also in T.events)
T.books = {};         // course id → { book: [...paras], packs: {chapter: pack} }
T.greek = { vocab: [], alphabet: [], concepts: [], quizzes: [] };
T.media = {};         // "courseId/name" → object URL
T.status = { text: "", busy: false, lastSync: null, offline: false, pending: false };
const listeners = new Set();
T.onChange = fn => listeners.add(fn);
const changed = () => listeners.forEach(fn => { try { fn(); } catch (e) { console.error(e); } });

const parse = (text, fallback) => { try { return text == null ? fallback : JSON.parse(text.replace(/^﻿/, "")); } catch { return undefined; } };

/** Loads the copy on this device (works offline), then syncs with OneDrive. */
T.start = async () => {
  const kv = async k => DB.get("kv", k);
  T.config = (await kv("config")) || T.config;
  T.mine = (await kv("mine")) || [];
  const others = (await kv("others")) || [];
  T.events = mergeEvents(others, T.mine);
  for (const g of ["vocab", "alphabet", "concepts", "quizzes"]) T.greek[g] = (await kv("greek-" + g)) || [];
  T.greekDirty = (await kv("greek-dirty")) || {};
  for (const c of T.config.Courses || []) await loadBook(c.Id);
  T.status.pending = T.mine.some(e => e._pending);
  changed();
  T.sync();
  setInterval(() => { if (!document.hidden) T.sync({ quick: true }); }, 60000);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) T.sync({ quick: true }); });
  window.addEventListener("online", () => T.sync({ quick: true }));
};

function mergeEvents(others, mine) {
  const all = others.filter(e => e.Device !== T.device).concat(mine);
  return all.sort((a, b) => (a.At < b.At ? -1 : a.At > b.At ? 1 : (a.Device || "").localeCompare(b.Device || "") || a.Seq - b.Seq));
}

async function loadBook(id) {
  const b = await DB.get("kv", "book:" + id);
  if (b) T.books[id] = b;
}

/** The newest from OneDrive: what changed since last time is downloaded and kept here for offline use. */
let syncing = null;
T.sync = (opts = {}) => syncing || (syncing = runSync(opts).finally(() => { syncing = null; }));
async function runSync({ quick = false } = {}) {
  if (typeof account !== "undefined" && !account && !window.TRAINING_REMOTE) return;
  T.status.busy = true; T.status.text = "Syncing…"; changed();
  try {
    await pushMine();
    await pushGreek();
    const etags = (await DB.get("kv", "etags")) || {};
    const fresh = async (path, etag, binary) => {
      if (etags[path] === etag) return null;
      const body = await T.remote.get(path, binary);
      return { body, done: () => { etags[path] = etag; } };
    };
    // The library and every device's progress.
    const top = await T.remote.list("training");
    const byName = Object.fromEntries(top.map(x => [x.name, x]));
    if (byName["course.json"]) {
      const f = await fresh("training/course.json", byName["course.json"].etag);
      if (f) { const cfg = parse(f.body); if (cfg) { T.config = cfg; await DB.put("kv", "config", cfg); f.done(); } }
    }
    let others = (await DB.get("kv", "others")) || [];
    const byDevice = {};
    for (const e of others) (byDevice[e.Device] ||= []).push(e);
    let logsChanged = false;
    for (const it of top.filter(x => /^progress-.*\.json$/i.test(x.name) && !/conflict/i.test(x.name))) {
      const path = "training/" + it.name;
      if (it.name === MY_LOG) {
        // Our own log on OneDrive: after clearing this device's storage, it brings your progress back.
        if (!T.mine.length) {
          const f = await fresh(path, it.etag);
          const list = f && parse(f.body, []);
          if (Array.isArray(list)) { T.mine = list; await DB.put("kv", "mine", T.mine); f.done(); logsChanged = true; }
        }
        continue;
      }
      const f = await fresh(path, it.etag);
      if (!f) continue;
      const list = parse(f.body, []);
      if (!Array.isArray(list)) continue; // OneDrive still writing it: next time
      const dev = list[0]?.Device || it.name;
      for (const k of Object.keys(byDevice)) if (k === dev) delete byDevice[k];
      byDevice[dev] = list;
      f.done(); logsChanged = true;
    }
    if (logsChanged) {
      others = Object.values(byDevice).flat();
      await DB.put("kv", "others", others);
      T.events = mergeEvents(others, T.mine);
    }
    // Each book: its text (book.json), chapter packs, pictures and cover.
    for (const c of T.config.Courses || []) {
      const dir = T.safe(c.Id);
      if (!/^[A-Za-z0-9_-]{1,80}$/.test(dir)) continue;
      const files = await T.remote.list("training/" + dir);
      const entry = T.books[c.Id] || { book: [], packs: {} };
      let bookChanged = false;
      for (const it of files) {
        const path = `training/${dir}/${it.name}`;
        if (it.name === "book.json" || /^ch\d{2,3}\.focuspack$/.test(it.name)) {
          const f = await fresh(path, it.etag);
          if (!f) continue;
          const data = parse(f.body);
          if (data === undefined) continue;
          if (it.name === "book.json") entry.book = data || [];
          else if (data && data.chapter) entry.packs[data.chapter] = data;
          f.done(); bookChanged = true;
        }
      }
      // Packs removed on the computer leave here too.
      const names = new Set(files.map(x => x.name));
      for (const ch of Object.keys(entry.packs)) if (!names.has(`ch${String(ch).padStart(2, "0")}.focuspack`) && !names.has(`ch${ch}.focuspack`)) { delete entry.packs[ch]; bookChanged = true; }
      if (bookChanged) { T.books[c.Id] = entry; await DB.put("kv", "book:" + c.Id, entry); bookCache.delete(c.Id); }
      else T.books[c.Id] = entry;
      if (!quick && files.some(x => x.folder && x.name === "media")) {
        const media = await T.remote.list(`training/${dir}/media`);
        for (const it of media) {
          const path = `training/${dir}/media/${it.name}`;
          if (etags[path] === it.etag && await DB.get("files", path)) continue;
          const blob = await T.remote.get(path, true);
          if (blob) { await DB.put("files", path, blob); etags[path] = it.etag; }
        }
      }
    }
    if (!quick && byName.covers) {
      for (const it of await T.remote.list("training/covers")) {
        const path = "training/covers/" + it.name;
        if (etags[path] === it.etag && await DB.get("files", path)) continue;
        const blob = await T.remote.get(path, true);
        if (blob) { await DB.put("files", path, blob); etags[path] = it.etag; }
      }
    }
    // Greek: your words, alphabet and concepts (shared with your computers), and the concept quiz packs.
    for (const g of ["vocab", "alphabet", "concepts"]) {
      const f = await T.remote.get(g + ".json");
      const list = parse(f, []);
      if (Array.isArray(list)) { T.greek[g] = mergeGreekIn(g, list); await DB.put("kv", "greek-" + g, T.greek[g]); }
    }
    if (!quick) {
      const quizzes = [];
      for (const it of (await T.remote.list("greek-concepts")).filter(x => /\.quizpack$/i.test(x.name))) {
        const p = parse(await T.remote.get("greek-concepts/" + it.name));
        if (p && (p.Questions || p.questions)) quizzes.push(normQuiz(p));
      }
      quizzes.sort((a, b) => lessonKeyCompare(a.Lesson, b.Lesson));
      T.greek.quizzes = quizzes; await DB.put("kv", "greek-quizzes", quizzes);
    }
    await DB.put("kv", "etags", etags);
    T.status.lastSync = new Date(); T.status.offline = false;
    T.status.text = "Up to date, " + T.status.lastSync.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  } catch (e) {
    T.status.offline = !navigator.onLine || /Failed to fetch|NetworkError/i.test(e.message);
    T.status.text = T.status.offline ? "Offline: your progress is kept here and syncs when you're back online." : e.message;
  } finally {
    T.status.busy = false; T.status.pending = T.mine.some(e => e._pending);
    changed();
  }
}

// ---------- Your progress (this device's own log) ----------

/** Adds something that happened (an answer, a finished lesson…), exactly as a computer would. */
T.add = (e, courseId) => {
  const seq = T.mine.reduce((m, x) => Math.max(m, x.Seq || 0), 0) + 1;
  const ev = {
    Device: T.device, CourseId: courseId || e.CourseId || "", Seq: seq, At: new Date().toISOString(), Day: e.Day || T.todayYmd(),
    Kind: "", Question: "", Concept: "", Correct: false, Sure: false, Lesson: -1, Note: "", Chapter: 0, Para: -1, Quote: "", Color: "", Grade: -1,
    ...e, _pending: true
  };
  if (!ev.CourseId) ev.CourseId = courseId || "";
  ev.Seq = seq;
  T.mine.push(ev); T.events.push(ev);
  DB.put("kv", "mine", T.mine);
  T.status.pending = true;
  schedulePush();
  invalidate();
  return ev;
};
let pushTimer = null;
function schedulePush() { clearTimeout(pushTimer); pushTimer = setTimeout(() => pushMine().then(changed, () => {}), 1500); }
async function pushMine() {
  if (!T.mine.some(e => e._pending)) return;
  const clean = T.mine.map(({ _pending, ...rest }) => rest);
  await T.remote.put("training/" + MY_LOG, JSON.stringify(clean));
  T.mine.forEach(e => delete e._pending);
  await DB.put("kv", "mine", T.mine);
  T.status.pending = false;
}

// ---------- Greek progress (shared files: only the cards you practised change) ----------

function normQuiz(p) {
  const g = (o, k) => o[k] ?? o[k[0].toLowerCase() + k.slice(1)];
  return {
    Lesson: String(g(p, "Lesson") || ""), Title: g(p, "Title") || "",
    Sections: (g(p, "Sections") || []).map(s => ({ Id: g(s, "Id") || "", Heading: g(s, "Heading") || "", Paragraphs: g(s, "Paragraphs") || [] })),
    Questions: (g(p, "Questions") || []).map(q => ({ Id: g(q, "Id") || "", Prompt: g(q, "Prompt") || "", Options: g(q, "Options") || [], Correct: g(q, "Correct") || [], Explanation: g(q, "Explanation") || "", Section: g(q, "Section") || "" }))
  };
}
function lessonKeyCompare(a, b) {
  const k = s => String(s).split(/[.\- ]/).map(x => parseFloat(x) || 0);
  const x = k(a), y = k(b);
  for (let i = 0; i < Math.max(x.length, y.length); i++) { const p = x[i] || 0, q = y[i] || 0; if (p !== q) return p - q; }
  return String(a).localeCompare(String(b));
}
T.lessonKeyCompare = lessonKeyCompare;

/** OneDrive's copy, with the cards you changed here (and haven't saved yet) kept as yours. */
function mergeGreekIn(kind, remote) {
  const dirty = T.greekDirty[kind] || {};
  const local = Object.fromEntries((T.greek[kind] || []).map(c => [c.Id, c]));
  return remote.map(r => {
    const at = dirty[r.Id];
    if (!at || !local[r.Id]) return r;
    const rm = T.date(r.Modified);
    return rm && rm > new Date(at) ? r : local[r.Id];
  });
}
/** Records a Greek card you practised; saved to OneDrive soon (merged with your computers' changes). */
T.greekChanged = (kind, card) => {
  const now = new Date();
  card.Modified = now.toISOString();
  (T.greekDirty[kind] ||= {})[card.Id] = now.toISOString();
  DB.put("kv", "greek-" + kind, T.greek[kind]);
  DB.put("kv", "greek-dirty", T.greekDirty);
  clearTimeout(T._greekTimer);
  T._greekTimer = setTimeout(() => pushGreek().then(changed, () => {}), 2000);
};
async function pushGreek() {
  for (const kind of ["vocab", "alphabet", "concepts"]) {
    const dirty = T.greekDirty[kind];
    if (!dirty || !Object.keys(dirty).length) continue;
    const text = await T.remote.get(kind + ".json");
    const remote = parse(text, []);
    if (!Array.isArray(remote)) continue; // busy: next time
    const local = Object.fromEntries(T.greek[kind].map(c => [c.Id, c]));
    let touched = false;
    const merged = remote.map(r => {
      const at = dirty[r.Id];
      if (!at || !local[r.Id]) return r;
      const rm = T.date(r.Modified);
      if (rm && rm > new Date(at)) return r; // changed on a computer after you practised it here
      touched = true;
      return { ...r, ...pick(local[r.Id], ["Level", "Due", "LastSeen", "Right", "Wrong", "LastWrong", "Modified"]) };
    });
    if (touched) await T.remote.put(kind + ".json", JSON.stringify(merged, null, 2));
    T.greek[kind] = merged;
    T.greekDirty[kind] = {};
    await DB.put("kv", "greek-" + kind, merged);
    await DB.put("kv", "greek-dirty", T.greekDirty);
  }
}
const pick = (o, keys) => Object.fromEntries(keys.filter(k => k in o).map(k => [k, o[k]]));

// ---------- Pictures and covers (kept on this device) ----------

T.mediaUrl = async (courseId, name) => {
  const key = `training/${T.safe(courseId)}/media/${name}`;
  if (T.media[key]) return T.media[key];
  const blob = await DB.get("files", key);
  if (!blob) return null;
  return (T.media[key] = URL.createObjectURL(blob));
};
T.coverUrl = async c => {
  if (!c.Cover) return null;
  const key = "training/covers/" + c.Cover;
  if (T.media[key]) return T.media[key];
  const blob = await DB.get("files", key);
  return blob ? (T.media[key] = URL.createObjectURL(blob)) : null;
};

// ======================= The course logic (the same as Focus on your computers) =======================

const CourseOf = e => (e.CourseId && e.CourseId.length ? e.CourseId : "desilva-nt");
T.courseOf = CourseOf;
T.course = id => (T.config.Courses || []).find(c => c.Id === id);

// Caches, cleared whenever something new is recorded or synced.
const bookCache = new Map();
let memo = new Map();
function invalidate() { memo = new Map(); }
T.onChange(invalidate);
const cached = (key, fn) => { if (memo.has(key)) return memo.get(key); const v = fn(); memo.set(key, v); return v; };

/** One book: its paragraphs and packs, lessons in order, and quick lookups. */
T.B = id => {
  if (bookCache.has(id)) return bookCache.get(id);
  const raw = T.books[id] || { book: [], packs: {} };
  const book = raw.book || [];
  const norm = book.map(p => normalize(paraText(p)));
  const packs = Object.values(raw.packs || {}).sort((a, b) => a.chapter - b.chapter);
  const b = { id, info: T.course(id) || { Id: id, Title: id, Chapters: 24 }, book, norm, packs, packOf: ch => packs.find(p => p.chapter === ch) };
  // Every imported lesson, chapter by chapter, stopping at the first chapter without a pack.
  b.lessons = [];
  let expected = packs.length ? packs[0].chapter : 1;
  for (const pack of packs) {
    if (pack.chapter !== expected) break;
    (pack.lessons || []).forEach((l, i) => b.lessons.push({ index: b.lessons.length, pack, lesson: fill(l), number: i + 1, course: id }));
    expected++;
  }
  b.cardIds = new Set(b.lessons.flatMap(l => l.lesson.cards.map(c => c.id)));
  bookCache.set(id, b);
  return b;
};
const fill = l => ({ title: "", bigQuestions: [], parts: [], warmup: [], practice: [], cards: [], quiz: [], timeline: [], mainIdeas: [], outline: [], endsChapter: false, ...l });
const paraText = p => (p.Runs || []).map(r => r.Text || "").join("");
T.paraText = paraText;
T.allQuestions = l => [...l.warmup, ...l.practice, ...l.cards, ...l.quiz];

function normalize(s) {
  let out = "", space = false;
  for (const c0 of String(s).toLowerCase()) {
    let c = c0;
    if ("‘’‛′".includes(c)) c = "'"; else if ("“”„″".includes(c)) c = '"'; else if ("–—‑‒".includes(c)) c = "-"; else if (c === "­") continue;
    if (/\s/.test(c)) { space = true; continue; }
    if (space && out.length) out += " ";
    space = false; out += c;
  }
  return out;
}
T.normalize = normalize;

function find(b, text, from = 0, to = Infinity, preferHeading = false) {
  const want = normalize(text);
  if (!want) return -1;
  to = Math.min(to, b.book.length);
  if (preferHeading) for (let i = Math.max(0, from); i < to; i++) if (b.book[i].Heading > 0 && b.norm[i].includes(want)) return i;
  for (let i = Math.max(0, from); i < to; i++) if (b.norm[i].includes(want)) return i;
  return -1;
}
function chapterRange(b, pack) {
  const start = find(b, pack.startText || "", 0, Infinity, true);
  if (start < 0) return [-1, -1];
  const end = pack.endBefore && pack.endBefore.trim() ? find(b, pack.endBefore, start + 1, Infinity, true) : -1;
  return [start, end < 0 ? b.book.length : end];
}
const STOP = new Set(["the", "and", "for", "from", "with", "into", "of", "a", "an", "to", "in", "on"]);
function findPart(b, l, part, from, to) {
  if (part < 0 || part >= l.parts.length) return -1;
  const p = l.parts[part];
  const exact = find(b, p.start || p.title || "", from, to, true);
  if (exact >= 0 && b.book[exact].Heading > 0) return exact;
  const words = normalize(p.title || p.start || "").split(/[ ,.:;"'()\-]+/).filter(w => w.length >= 3 && !STOP.has(w));
  if (words.length) for (let k = Math.max(0, from); k < Math.min(to, b.book.length); k++)
    if (b.book[k].Heading > 0 && paraText(b.book[k]).length < 200 && words.every(w => b.norm[k].includes(w))) return k;
  if (exact >= 0) return exact;
  const refs = T.allQuestions(l).flatMap(q => q.refs || []).filter(r => r.part === part).map(r => find(b, r.find, from, to)).filter(x => x >= 0);
  return refs.length ? Math.min(...refs) : -1;
}
T.chapterRange = (id, pack) => chapterRange(T.B(id), pack);
/** The paragraphs a lesson reads: from its first part to the next lesson's first part (or the chapter's end). */
T.lessonRange = lr => {
  const b = T.B(lr.course);
  const [cs, ce] = chapterRange(b, lr.pack);
  if (cs < 0) return [-1, -1];
  const idx = lr.number - 1;
  const first = (l, from, to) => (l.parts.length ? (x => (x < 0 ? null : x))(findPart(b, l, 0, from, to)) : null);
  const start = idx === 0 ? cs : first(lr.lesson, cs, ce) ?? cs;
  let end = ce;
  if (idx + 1 < lr.pack.lessons.length) { const nx = first(fill(lr.pack.lessons[idx + 1]), start + 1, ce); if (nx != null) end = nx; }
  return [start, end];
};
/** Where an answer is in the book: a paragraph index (searched within the lesson's chapter). */
T.locate = (lr, ref) => {
  const b = T.B(lr.course);
  const [cs, ce] = chapterRange(b, lr.pack);
  if (cs < 0) return -1;
  let from = cs;
  if (ref.part >= 0 && ref.part < lr.lesson.parts.length) { const p = findPart(b, lr.lesson, ref.part, cs, ce); if (p >= 0) from = p; }
  const i = find(b, ref.find || "", from, ce);
  return i >= 0 ? i : find(b, ref.find || "", cs, ce);
};
T.findInBook = (courseId, text, from, to) => find(T.B(courseId), text, from, to);

// ----- Events for one book -----
T.CE = id => cached("ce:" + id, () => T.events.filter(e => CourseOf(e) === id));
T.doneLessons = id => cached("done:" + id, () => new Set(T.CE(id).filter(e => e.Kind === "lessonDone" && e.Lesson >= 0).map(e => e.Lesson)));
T.stepsDone = (id, lesson) => new Set(T.CE(id).filter(e => e.Kind === "lessonStep" && e.Lesson === lesson).map(e => e.Question));
T.nextLesson = id => { const done = T.doneLessons(id); return T.B(id).lessons.find(l => !done.has(l.index)) || null; };
T.inProgress = lr => !!lr && !T.doneLessons(lr.course).has(lr.index) && T.stepsDone(lr.course, lr.index).size > 0;
T.hidden = id => cached("hidden:" + id, () => {
  const set = new Set();
  for (const e of T.CE(id)) { if (e.Kind === "hide") set.add(e.Question); else if (e.Kind === "unhide") set.delete(e.Question); }
  for (const [q, choice] of Object.entries(T.cardSorts(id))) if (choice === "retire") set.add(q);
  return set;
});
T.cardSorts = id => cached("sorts:" + id, () => { const d = {}; for (const e of T.CE(id)) if (e.Kind === "cardSort") d[e.Question] = e.Note; return d; });
T.voteOf = (id, q) => { const v = T.CE(id).filter(e => e.Kind === "vote" && e.Question === q).pop(); return v ? v.Grade : 0; };
T.isPrioritized = (id, q) => { const v = T.CE(id).filter(e => e.Kind === "priority" && e.Question === q).pop(); return !!v && v.Grade > 0; };
T.isChapterCard = (id, q) => cached("chcards:" + id, () => {
  const sorts = T.cardSorts(id);
  return new Set(T.B(id).lessons.flatMap(l => l.lesson.cards).filter(c => c.tier !== "core" && sorts[c.id] !== "keep").map(c => c.id));
}).has(q);
T.chaptersPassed = id => cached("passed:" + id, () => {
  const first = {};
  for (const e of T.CE(id)) if (e.Kind === "answer" && !(e.Question in first)) first[e.Question] = e.Grade >= 1 || (e.Grade < 0 && e.Correct);
  const passed = new Set();
  for (const pack of T.B(id).packs) {
    const quiz = (pack.lessons || []).flatMap(l => (l.quiz || []).map(q => q.id));
    if (!quiz.length || !quiz.every(q => q in first)) continue;
    if (quiz.filter(q => first[q]).length >= 0.8 * quiz.length) passed.add(pack.chapter);
  }
  return passed;
});
T.cardsOff = id => cached("off:" + id, () => {
  const off = new Set(Object.entries(T.cardSorts(id)).filter(([, v]) => v === "quiz" || v === "retire").map(([k]) => k));
  const passed = T.chaptersPassed(id);
  for (const l of T.B(id).lessons.filter(l => passed.has(l.pack.chapter))) for (const c of l.lesson.cards) if (T.isChapterCard(id, c.id)) off.add(c.id);
  return off;
});

/** A question's review state from every answer on every device (the same rules as Focus on your computers). */
T.stateOf = (id, q, longTerm) => {
  if (longTerm === undefined) longTerm = T.B(id).cardIds.has(q);
  return cached(`st:${id}:${q}:${longTerm}`, () => {
    let due = null, streak = 0, seen = 0, miscon = false, last = null;
    const gaps = longTerm ? (T.isChapterCard(id, q) ? [1, 3, 7, 14] : [1, 3, 7, 14, 30, 60, 120]) : [3];
    for (const e of T.CE(id)) {
      if (e.Kind !== "answer" || e.Question !== q) continue;
      seen++;
      const day = T.parseDay(e.Day) || T.startOfDay(new Date(e.At));
      const lasted = last ? Math.max(0, Math.round((day - last) / DAY)) : 0;
      const late = due && day > due;
      last = day;
      if (e.Grade === 0) { streak = 0; miscon = false; due = T.addDays(day, 1); }
      else if (e.Grade === 2) {
        streak = Math.max(streak + 2, 3); miscon = false;
        let g = gaps[Math.min(streak - 1, gaps.length - 1)] * 2;
        if (late) g = Math.max(g, Math.min(lasted * 2, 365));
        due = T.addDays(day, g);
      } else if (e.Grade === 1 || (e.Correct && e.Sure)) {
        streak++; miscon = false;
        let g = gaps[Math.min(streak - 1, gaps.length - 1)];
        if (late) g = Math.max(g, Math.min(lasted * 2, 365));
        due = T.addDays(day, g);
      } else { streak = 0; miscon = !e.Correct && e.Sure; due = T.addDays(day, 1); }
    }
    const vote = T.voteOf(id, q);
    if (vote !== 0 && last && due && due > last) { const gap = Math.round((due - last) / DAY); due = T.addDays(last, vote > 0 ? Math.max(1, Math.floor(gap / 2)) : gap * 2); }
    return { due: due || new Date(0), streak, misconception: miscon, seen, graduated: streak > gaps.length };
  });
};
T.released = (id, q) => {
  if (T.CE(id).some(e => e.Kind === "release" && e.Question === q)) return true;
  const first = T.CE(id).find(e => e.Kind === "answer" && e.Question === q);
  const since = T.date(T.config.IntakeSinceUtc);
  return !!first && !!since && new Date(first.At) < since;
};
T.retiredConcepts = id => cached("ret:" + id, () => {
  const groups = {};
  for (const e of T.CE(id)) if (e.Kind === "answer" && e.Concept && e.Correct && e.Sure) (groups[e.Concept] ||= new Set()).add(e.Day || T.ymd(new Date(e.At)));
  const out = new Set();
  for (const [c, days] of Object.entries(groups)) { const d = [...days].sort(); if (d.length >= 6 && (T.parseDay(d[d.length - 1]) - T.parseDay(d[0])) / DAY >= 90) out.add(c); }
  return out;
});
const reviewables = l => [...l.lesson.practice.map(q => [q, false]), ...l.lesson.cards.map(q => [q, true]), ...l.lesson.quiz.filter(q => q.kind === "mc").map(q => [q, false])];
/** Lets waiting memory cards into the review while today's review fits your daily review time. */
T.releaseCards = id => {
  if (!T.config.IntakeSinceUtc) return; // set up on your computer
  const hidden = T.hidden(id), done = T.doneLessons(id), retired = T.retiredConcepts(id), off = T.cardsOff(id), today = T.startOfDay(new Date());
  let load = 0;
  const waiting = [];
  for (const l of T.B(id).lessons.filter(l => done.has(l.index)))
    for (const [q, card] of reviewables(l)) {
      if (hidden.has(q.id) || (card && off.has(q.id)) || (q.concept && retired.has(q.concept))) continue;
      const st = T.stateOf(id, q.id, card);
      if (st.seen === 0 || st.graduated) continue;
      if (card && !T.released(id, q.id)) { waiting.push(q.id); continue; }
      if (st.due <= today) load += card ? 12 : 25;
    }
  let room = Math.max(1, T.config.ReviewMinutes || 10) * 60 - load;
  for (const q of waiting) { if (room < 12) break; T.add({ Kind: "release", Question: q }, id); room -= 12; }
};
/** Review for today: due questions from lessons already done (this book only), misconceptions first. */
T.dueReview = (id, cap = 999) => {
  T.releaseCards(id);
  return cached(`due:${id}:${cap}`, () => {
    const hidden = T.hidden(id), done = T.doneLessons(id), retired = T.retiredConcepts(id), off = T.cardsOff(id), today = T.startOfDay(new Date());
    const due = [];
    for (const l of T.B(id).lessons.filter(l => done.has(l.index)))
      for (const [q, card] of reviewables(l)) {
        if (hidden.has(q.id) || (card && off.has(q.id))) continue;
        if (q.concept && retired.has(q.concept)) continue;
        if (card && !T.released(id, q.id)) continue;
        const s = T.stateOf(id, q.id, card);
        if (s.graduated) continue;
        if (s.seen > 0 && s.due <= today) due.push({ l, q, s });
      }
    due.sort((a, b) => (b.s.misconception - a.s.misconception) || (a.s.due - b.s.due) || (a.s.streak - b.s.streak));
    return due.slice(0, cap).map(x => ({ l: x.l, q: x.q }));
  });
};
/** Today's lesson cards, hardest first (for Book flashcards when nothing is due). */
T.hardestCardsToday = (id, max = 40) => {
  const today = T.todayYmd();
  const idx = new Set(T.CE(id).filter(e => e.Day === today && e.Lesson >= 0 && (e.Kind === "lessonDone" || e.Kind === "lessonStep")).map(e => e.Lesson));
  if (!idx.size) return [];
  const hidden = T.hidden(id), off = T.cardsOff(id);
  const answers = {};
  for (const e of T.CE(id)) if (e.Kind === "answer" && e.Day === today) (answers[e.Question] ||= []).push(e);
  const list = [];
  for (const l of T.B(id).lessons.filter(l => idx.has(l.index)))
    for (const q of l.lesson.cards) {
      const a = answers[q.id];
      if (hidden.has(q.id) || off.has(q.id) || !a) continue;
      const misses = a.filter(x => x.Grade === 0 || (x.Grade < 0 && !x.Correct)).length;
      const last = a[a.length - 1], lastGrade = last.Grade >= 0 ? last.Grade : last.Correct ? 1 : 0;
      const pri = T.isPrioritized(id, q.id), vote = T.voteOf(id, q.id);
      if (lastGrade === 2 && misses === 0 && !pri && vote <= 0) continue;
      list.push({ l, q, score: misses * 2 + (lastGrade === 0 ? 2 : lastGrade === 1 ? 1 : 0) + (pri ? 100 : 0) + vote * 100 });
    }
  return list.sort((a, b) => b.score - a.score).slice(0, Math.max(1, max)).map(x => ({ l: x.l, q: x.q }));
};
T.bookFraction = id => {
  const b = T.B(id), done = T.doneLessons(id);
  let f = 0;
  for (const pack of b.packs) { const mine = b.lessons.filter(l => l.pack.chapter === pack.chapter); if (mine.length) f += mine.filter(l => done.has(l.index)).length / mine.length; }
  return Math.min(1, Math.max(0, f / Math.max(1, b.info.Chapters || 24)));
};

// ----- Which book each day (the weekly plan, changed days and days off) -----
const ROLES = ["Track A", "Track B", "Elective"];
T.dayOff = d => (T.config.DaysOff || {})[T.ymd(d)] || "";
function planBook(d) {
  const dow = d.getDay();
  if (dow === 0 || dow === 6) return null;
  const cs = T.config.Courses || [];
  if (cs.every(c => !(c.Days || []).length)) return cs.find(c => c.Id === T.config.ActiveId) || cs[0] || null;
  return cs.filter(c => (c.Days || []).includes(dow)).sort((a, b) => ROLES.indexOf(a.Role || "") - ROLES.indexOf(b.Role || ""))[0] || null;
}
T.booksFor = d => {
  if (T.dayOff(d)) return [];
  const plan = (T.config.WeekPlan || {})[T.ymd(d)];
  if (plan && !plan.Default) return (plan.Books || []).map(T.course).filter(Boolean);
  const c = planBook(d);
  return c ? [c] : [];
};
T.doneCountOn = (id, d) => { const day = T.ymd(d); return T.CE(id).filter(e => e.Kind === "lessonDone" && e.Day === day).length; };
T.dayLessons = d => {
  const seen = {};
  return T.booksFor(d).map(b => { const n = seen[b.Id] = (seen[b.Id] || 0) + 1; return { book: b, done: T.doneCountOn(b.Id, d) >= n }; });
};
T.nextDueToday = () => (T.dayLessons(new Date()).find(x => !x.done) || {}).book || null;
/** The last weekday's planned lesson, if it wasn't done and hasn't been caught up (an optional catch-up). */
T.missedLesson = () => {
  if ((T.config.MissedPolicy || "offer") !== "offer") return null;
  const today = T.startOfDay(new Date());
  for (let d = T.addDays(today, -1); (today - d) / DAY <= 3; d = T.addDays(d, -1)) {
    const lessons = T.dayLessons(d);
    if (!lessons.length) continue;
    const miss = lessons.find(x => !x.done);
    if (!miss) return null;
    if ((T.config.SkippedCatchUps || []).includes(miss.book.Id + "|" + T.ymd(d))) return null;
    for (let x = T.addDays(d, 1); x <= today; x = T.addDays(x, 1))
      if (T.doneCountOn(miss.book.Id, x) > T.booksFor(x).filter(c => c.Id === miss.book.Id).length) return null;
    return { book: miss.book, day: d };
  }
  return null;
};

// ----- Timeline (spaced by event, same rules as the computers) -----
T.timelineState = (bookId, ev) => {
  let due = new Date(0), gap = 0, streak = 0, seen = 0;
  for (const e of T.events) {
    if (e.Kind !== "timeline" || e.Question !== ev.text || CourseOf(e) !== bookId) continue;
    seen++;
    const at = new Date(e.At);
    const off = e.Grade >= 0 ? e.Grade : e.Correct ? 0 : 999;
    if (e.Correct) {
      const exact = off === 0;
      streak = exact ? streak + 1 : 0;
      gap = gap <= 0 ? (exact ? 2 : 1) : gap * (exact ? 2.5 : 1.4);
      due = new Date(at.getTime() + gap * DAY);
    } else {
      streak = 0; gap = 0;
      const far = off / Math.max(5, ev.within ?? 10);
      due = new Date(at.getTime() + (far > 8 ? 600000 : far > 3 ? 3 * 3600000 : DAY));
    }
  }
  return { due, streak, seen, graduated: streak >= 4 };
};
T.learnedEvents = ids => {
  const list = [];
  for (const id of ids) {
    const done = T.doneLessons(id);
    let order = 0;
    for (const l of T.B(id).lessons.filter(l => done.has(l.index))) { order++; for (const ev of l.lesson.timeline) list.push({ l, ev, book: id, order }); }
  }
  return list;
};
T.timelineRustiest = (ids, max) => {
  const now = new Date();
  return T.learnedEvents(ids).map(x => {
    const st = T.timelineState(x.book, x.ev);
    const key = st.seen === 0 ? -1e6 + Math.random() : st.due <= now && !st.graduated ? -1e7 + st.due.getTime() / 1e14 : ((st.due - now) / DAY) * (0.6 + Math.random() * 0.8) + (st.graduated ? 30 : 0);
    return { x, key };
  }).sort((a, b) => a.key - b.key).slice(0, max).map(a => a.x);
};
T.timelineDue = ids => T.learnedEvents(ids).filter(x => { const st = T.timelineState(x.book, x.ev); return st.seen > 0 && !st.graduated && st.due <= new Date(); }).length;
T.timelineFor = (lr, max = 8) => {
  const done = T.doneLessons(lr.course);
  const misses = text => T.CE(lr.course).filter(e => e.Kind === "timeline" && e.Question === text).reduce((n, e) => n + (e.Correct ? -1 : 1), 0);
  const mine = lr.lesson.timeline.map(ev => ({ l: lr, ev, book: lr.course }));
  const earlier = T.shuffle(T.B(lr.course).lessons.filter(l => l.index !== lr.index && done.has(l.index)).flatMap(l => l.lesson.timeline.map(ev => ({ l, ev, book: lr.course }))))
    .sort((a, b) => misses(b.ev.text) - misses(a.ev.text));
  return mine.concat(earlier).slice(0, Math.max(mine.length, max));
};
T.yearLabel = y => (y < 0 ? `${-y} BCE` : `${Math.max(1, y)} CE`);
/** A year typed as 4 BC, 4 BCE, -4, 30, 30 AD, AD 30, c. 30 or 30 CE (null if it isn't one). */
T.parseYear = text => {
  let t = String(text).trim().toUpperCase().replace(/[.,]/g, "");
  if (t.startsWith("C ") || t.startsWith("CA ")) t = t.slice(t.indexOf(" ") + 1);
  const bce = /B\s*C\s*E?$|^-/.test(t);
  t = t.replace(/B\s*C\s*E?|C\s*E|A\s*D|-|\s/g, "");
  if (!/^\d{1,4}$/.test(t)) return null;
  const y = parseInt(t, 10);
  return bce ? -Math.abs(y) : y;
};

// ----- Exams and finished books -----
T.examsPassed = id => new Set(T.CE(id).filter(e => e.Kind === "examDone" && e.Correct).map(e => e.Note));
T.finished = id => T.examsPassed(id).has("final");
T.examDue = id => {
  const b = T.B(id), done = T.doneLessons(id), passed = T.examsPassed(id), lessons = b.lessons;
  const has = ch => b.packs.some(p => p.chapter === ch);
  const ends = (b.info.UnitEnds || []).length ? b.info.UnitEnds : Array.from({ length: Math.floor((b.info.Chapters || 24) / 3) }, (_, u) => (u + 1) * 3);
  let prev = 0;
  for (let u = 0; u < ends.length; u++) {
    const last = ends[u];
    const unit = lessons.filter(l => l.pack.chapter > prev && l.pack.chapter <= last);
    const complete = Array.from({ length: Math.max(0, last - prev) }, (_, i) => prev + 1 + i).every(has) && unit.length && unit.every(l => done.has(l.index));
    const key = `unit${u + 1}`;
    if (complete && !passed.has(key)) {
      const qs = examQuestions(id, unit, 16);
      if (prev > 0) qs.push(...examQuestions(id, lessons.filter(l => l.pack.chapter <= prev), 4));
      return { key, title: `Unit exam ${u + 1}: chapters ${prev + 1}–${last}` + (prev > 0 ? ", with a few from earlier" : ""), questions: qs };
    }
    prev = last;
  }
  const all = Array.from({ length: b.info.Chapters || 24 }, (_, i) => i + 1).every(has) && lessons.length && lessons.every(l => done.has(l.index));
  if (all && !passed.has("final")) return { key: "final", title: "Final exam: the whole book", questions: examQuestions(id, lessons, 40) };
  return null;
};
function examQuestions(id, lessons, count) {
  const hidden = T.hidden(id);
  return T.shuffle(lessons.flatMap(l => [...l.lesson.practice, ...l.lesson.quiz].filter(q => q.kind === "mc" && !hidden.has(q.id)).map(q => ({ l, q }))))
    .map(x => ({ ...x, s: T.stateOf(id, x.q.id) }))
    .sort((a, b) => (a.s.seen > 0) - (b.s.seen > 0) || a.s.due - b.s.due).slice(0, count).map(x => ({ l: x.l, q: x.q }));
}
T.allBooksRound = count => {
  const pool = [];
  for (const c of T.config.Courses || []) {
    if (!T.finished(c.Id)) continue;
    const hidden = T.hidden(c.Id);
    for (const l of T.B(c.Id).lessons) for (const q of T.allQuestions(l.lesson)) {
      if (q.kind !== "mc" || hidden.has(q.id)) continue;
      const st = T.stateOf(c.Id, q.id);
      const w = 1 + (st.misconception ? 3 : 0) + (st.streak === 0 ? 1.5 : 0) + Math.min(3, (Date.now() - st.due) / DAY / 30);
      pool.push({ l, q, key: -Math.log(Math.max(1e-9, Math.random())) / w });
    }
  }
  return pool.sort((a, b) => a.key - b.key).slice(0, count).map(x => ({ l: x.l, q: x.q }));
};

// ----- Highlights -----
T.highlights = id => {
  const map = new Map();
  for (const e of T.CE(id)) {
    if (e.Kind === "highlight") map.set(e.Question, { id: e.Question, chapter: e.Chapter, para: e.Para, quote: e.Quote, color: e.Color, note: e.Note });
    else if (e.Kind === "highlightNote" && map.has(e.Question)) map.get(e.Question).note = e.Note;
    else if (e.Kind === "unhighlight") map.delete(e.Question);
  }
  return [...map.values()];
};

// ----- Greek -----
const plain = s => String(s || "").normalize("NFD").replace(/\p{Mn}/gu, "").normalize("NFC").toLowerCase().replace(/ς/g, "σ");
T.plain = plain;
const missRate = c => ((c.Wrong || 0) + 0.5) / ((c.Right || 0) + (c.Wrong || 0) + 1);
T.greekDue = kind => {
  const now = new Date();
  return (T.greek[kind] || []).filter(c => (kind !== "vocab" || c.InReview !== false) && (T.date(c.Due) || now) <= now).length;
};
/** A round of Greek matching (the computers' rules): just missed, then due (least known first), then a draw favoring misses. */
T.greekRound = (kind, count) => {
  const now = new Date();
  const cards = (T.greek[kind] || []).filter(c => (kind !== "vocab" || c.InReview !== false) && (c.English || "").trim() && (c.Greek || "").trim());
  const due = c => (T.date(c.Due) || now) <= now;
  const missed = cards.filter(c => c.LastWrong);
  const dueNow = cards.filter(c => !c.LastWrong && due(c)).sort((a, b) => (a.Level || 0) - (b.Level || 0) || missRate(b) - missRate(a));
  const rest = cards.filter(c => !c.LastWrong && !due(c)).map(c => ({ c, k: -Math.log(1 - Math.random()) / (1 + 6 * missRate(c)) })).sort((a, b) => a.k - b.k).map(x => x.c);
  const picked = [];
  for (const c of [...missed, ...dueNow, ...rest]) {
    if (picked.length >= count) break;
    if (picked.some(p => plain(p.English).trim() === plain(c.English).trim() || plain(p.Greek) === plain(c.Greek))) continue;
    picked.push(c);
  }
  return T.shuffle(picked);
};
T.greekExtras = (kind, round, n) => {
  const out = [];
  for (const c of T.shuffle((T.greek[kind] || []).filter(c => !round.includes(c) && (c.English || "").trim()))) {
    if (out.length >= n) break;
    if ([...round, ...out].some(r => plain(r.English).trim() === plain(c.English).trim())) continue;
    out.push(c);
  }
  return out;
};
const H = 1 / 24;
const WAITS = {
  vocab: [2 * H, 1, 3, 7, 14, 21, 30, 40, 50, 60, 75],
  alphabet: [1 * H, 6 * H, 1, 3, 7, 14, 21, 30, 30],
  concepts: [2 * H, 8 * H, 1, 3, 7, 14, 21, 30, 30],
};
/** Records a match, by the same rules as your computers: right first time moves the card up a level
 *  (waiting longer, less if you often miss it); a miss drops it two levels and brings it back soon. */
T.greekRecord = (kind, c, right) => {
  const now = new Date();
  c.LastSeen = T.localIso(now);
  c.LastWrong = !right;
  if (right) {
    c.Right = (c.Right || 0) + 1;
    const waits = WAITS[kind];
    const max = kind === "vocab" ? (T.config.VocabRetiredLevel || 6) : 8;
    c.Level = Math.min(max, (c.Level || 0) + 1);
    const ease = kind === "alphabet" ? 1 : Math.min(1, Math.max(0.2, 1 / (1 + 5 * missRate(c) ** 2)));
    let days = waits[Math.min(c.Level, waits.length - 1)] * ease;
    if (kind === "vocab") days = Math.min(30, days);
    c.Due = T.localIso(new Date(now.getTime() + days * DAY));
  } else {
    c.Wrong = (c.Wrong || 0) + 1;
    c.Level = Math.max(0, (c.Level || 0) - 2);
    c.Due = T.localIso(new Date(now.getTime() + (kind === "alphabet" ? 20 : 30) * 60000));
  }
  T.greekChanged(kind, c);
};
