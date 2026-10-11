// Focus Tasks: task logic (mirrors Focus on your computer). No screen code, no network.
"use strict";
const DAY_NAMES = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
const MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
const ORDINALS = ["first", "second", "third", "fourth"];

function startOfDay(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function addMonths(d, n) { const x = new Date(d); x.setDate(1); x.setMonth(x.getMonth() + n); return x; }
function sameDay(a, b) { return !!a && !!b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }
function daysInMonth(y, m) { return new Date(y, m + 1, 0).getDate(); }
const pad = n => String(n).padStart(2, "0");
/** Local time with no time zone, the way Focus saves dates. */
function iso(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; }
function ymd(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function parseDate(s) { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; }
function parseYmd(s) { if (!s || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return null; const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); }
function words(text) { return text.toLowerCase().replace(/,/g, " , ").split(/\s+/).filter(Boolean); }

function weekday(word) {
  word = (word || "").replace(/[,.]/g, "").toLowerCase();
  if (word.length < 3) return null;
  for (let i = 0; i < 7; i++) { const n = DAY_NAMES[i]; if (word === n || word === n + "s" || n.startsWith(word)) return i; }
  return null;
}
function month(word) {
  word = (word || "").replace(/[,.]/g, "").toLowerCase();
  if (word.length < 3) return null;
  for (let i = 0; i < 12; i++) if (MONTHS[i] === word || MONTHS[i].startsWith(word)) return i;
  return null;
}
function dayNumber(word) {
  const m = /^(\d{1,2})(st|nd|rd|th)?$/.exec(word || "");
  if (!m) return null; const n = +m[1]; return n >= 1 && n <= 31 ? n : null;
}
function ordinal(word) {
  if (word === "last") return -1;
  const i = ORDINALS.indexOf(word); if (i >= 0) return i + 1;
  const o = { "1st": 1, "2nd": 2, "3rd": 3, "4th": 4 }[word];
  return o === undefined ? null : o;
}
function upcoming(today, m, d) {
  const y = today.getFullYear();
  let date = new Date(y, m, Math.min(d, daysInMonth(y, m)));
  if (date < addDays(today, -30)) date = new Date(y + 1, m, Math.min(d, daysInMonth(y + 1, m)));
  return date;
}
function nextWeekday(today, dow, strict) { let ahead = (dow - today.getDay() + 7) % 7; if (strict && ahead === 0) ahead = 7; return addDays(today, ahead); }
function endOfWeek(today) { return addDays(today, (5 - today.getDay() + 7) % 7); }

/** A date at w[i]: today, tomorrow, friday, next friday, next week, in 3 days, oct 14, 14 october, 10/14, end of week. */
function tryDate(w, i, today, inText) {
  const W = k => (w[k] || "").replace(/[,.]/g, "");
  const x = W(i);
  if (x === "today" || x === "tod" || x === "tonight") return { date: today, used: 1 };
  if (x === "tomorrow" || x === "tmr" || x === "tom") return { date: addDays(today, 1), used: 1 };
  if (x === "eow") return { date: endOfWeek(today), used: 1 };
  if (x === "eom") return { date: new Date(today.getFullYear(), today.getMonth() + 1, 0), used: 1 };
  if (x === "end" && W(i + 1) === "of") {
    if (W(i + 2) === "week") return { date: endOfWeek(today), used: 3 };
    if (W(i + 2) === "month") return { date: new Date(today.getFullYear(), today.getMonth() + 1, 0), used: 3 };
  }
  if (x === "next") {
    if (W(i + 1) === "week") return { date: addDays(today, 7 - (today.getDay() + 6) % 7), used: 2 };
    if (W(i + 1) === "month") return { date: new Date(today.getFullYear(), today.getMonth() + 1, 1), used: 2 };
    const d = weekday(W(i + 1)); if (d !== null) return { date: nextWeekday(today, d, true), used: 2 };
  }
  if (x === "in" && /^\d+$/.test(W(i + 1))) {
    const n = +W(i + 1), u = W(i + 2);
    if (n > 0 && n < 400) {
      if (u === "day" || u === "days") return { date: addDays(today, n), used: 3 };
      if (u === "week" || u === "weeks") return { date: addDays(today, 7 * n), used: 3 };
      if (u === "month" || u === "months") { const d = new Date(today); d.setMonth(d.getMonth() + n); return { date: d, used: 3 }; }
    }
  }
  const wd = weekday(x);
  if (wd !== null && !x.endsWith("s") && !(inText && (x === "sat" || x === "sun"))) return { date: nextWeekday(today, wd, false), used: 1 };
  const m1 = month(x), d1 = dayNumber(W(i + 1)); if (m1 !== null && d1 !== null) return { date: upcoming(today, m1, d1), used: 2 };
  const d2 = dayNumber(x), m2 = month(W(i + 1)); if (d2 !== null && m2 !== null) return { date: upcoming(today, m2, d2), used: 2 };
  const slash = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/.exec(x);
  if (slash) {
    const mm = +slash[1] - 1, dd = +slash[2];
    if (mm >= 0 && mm < 12 && dd >= 1 && dd <= 31) {
      if (slash[3]) { let y = +slash[3]; if (y < 100) y += 2000; return { date: new Date(y, mm, dd), used: 1 }; }
      return { date: upcoming(today, mm, dd), used: 1 };
    }
  }
  return null;
}

/** A time at w[i]: 3pm, 3:30pm, "at 15:30", "at 3" (afternoon), "at noon". */
function tryTime(w, i) {
  const W = k => (w[k] || "").replace(/[,.]/g, "");
  let k = i; const at = W(k) === "at"; if (at) k++;
  const x = W(k);
  if (x === "noon") return { h: 12, m: 0, used: k - i + 1 };
  const spaced = W(k + 1) === "am" || W(k + 1) === "pm";
  const text = spaced ? x + W(k + 1) : x;
  const pm = text.endsWith("pm"), am = text.endsWith("am");
  const core = pm || am ? text.slice(0, -2) : text;
  const m = /^(\d{1,2})(?::(\d{2}))?$/.exec(core);
  if (!m) return null;
  let h = +m[1]; const min = m[2] ? +m[2] : 0;
  if (!pm && !am && !at) return null; // a bare number ("call 3 people") stays as text
  if (min > 59) return null;
  if (pm && h < 12) h += 12;
  if (am && h === 12) h = 0;
  if (!pm && !am && !m[2] && h >= 1 && h <= 7) h += 12;
  if (h > 23) return null;
  return { h, m: min, used: k - i + (spaced ? 2 : 1) };
}

// ---------- Repeats (same rules as Focus) ----------

function parseRepeatAt(w, start, anchor) {
  const W = k => w[k] || "";
  const r = { type: "", interval: 1, days: [], monthDay: 0, ordinal: 0, weekday: 0, weekdaysOnly: false };
  let i = start;
  switch (W(i)) {
    case "daily": r.type = "days"; return { rule: r, used: 1 };
    case "weekdays": case "workdays": r.type = "days"; r.weekdaysOnly = true; return { rule: r, used: 1 };
    case "weekly": r.type = "weeks"; r.days = [anchor.getDay()]; i++; break;
    case "biweekly": r.type = "weeks"; r.interval = 2; r.days = [anchor.getDay()]; i++; break;
    case "monthly": r.type = "monthDay"; r.monthDay = anchor.getDate(); i++; break;
    case "yearly": case "annually": r.type = "years"; return { rule: r, used: 1 };
    case "every": case "each": i++; break;
    default: return null;
  }
  if (W(start) === "every" || W(start) === "each") {
    if (W(i) === "other") { r.interval = 2; i++; }
    else if (/^\d+$/.test(W(i)) && +W(i) > 0 && +W(i) < 100) { r.interval = +W(i); i++; }
    const unit = W(i);
    if (unit === "day" || unit === "days") { r.type = "days"; i++; }
    else if (["weekday", "weekdays", "workday", "workdays"].includes(unit)) { r.type = "days"; r.weekdaysOnly = true; i++; }
    else if (unit === "week" || unit === "weeks") { r.type = "weeks"; i++; r.days = [anchor.getDay()]; }
    else if (unit === "month" || unit === "months") { r.type = "monthDay"; r.monthDay = anchor.getDate(); i++; }
    else if (unit === "year" || unit === "years") { r.type = "years"; i++; }
    else if (ordinal(unit) !== null) {
      const ord = ordinal(unit), after = W(i + 1), dow = weekday(after);
      if (dow !== null) { r.type = "monthWeekday"; r.ordinal = ord; r.weekday = dow; i += 2; }
      else if (ord === -1 && (after === "weekday" || after === "workday")) { r.type = "monthLastWeekday"; i += 2; }
      else if (ord === -1 && after === "day") { r.type = "monthDay"; r.monthDay = -1; i += 2; }
      else return null;
      return { rule: r, used: i - start };
    }
    else if (dayNumber(unit) !== null) { r.type = "monthDay"; r.monthDay = dayNumber(unit); i++; return { rule: r, used: i - start }; }
    else if (weekday(unit) !== null) { r.type = "weeks"; }
    else return null;
  }
  if (r.type === "weeks") {
    let j = W(i) === "on" ? i + 1 : i; const days = [];
    while (weekday(W(j)) !== null) {
      days.push(weekday(W(j))); j++;
      if ((W(j) === "," || W(j) === "and") && weekday(W(j + 1)) !== null) j++;
    }
    if (days.length) { r.days = [...new Set(days)]; i = j; }
  } else if (r.type === "monthDay" && W(i) === "on") {
    let j = i + 1; if (W(j) === "the") j++;
    if (dayNumber(W(j)) !== null) { r.monthDay = dayNumber(W(j)); i = j + 1; }
    else if (ordinal(W(j)) !== null) {
      const ord = ordinal(W(j)), dow = weekday(W(j + 1));
      if (dow !== null) { r.type = "monthWeekday"; r.ordinal = ord; r.weekday = dow; i = j + 2; }
      else if (ord === -1 && (W(j + 1) === "weekday" || W(j + 1) === "workday")) { r.type = "monthLastWeekday"; i = j + 2; }
      else if (ord === -1 && W(j + 1) === "day") { r.monthDay = -1; i = j + 2; }
    }
  }
  return { rule: r, used: i - start };
}
function parseRepeat(text, anchor) {
  if (!text || !text.trim()) return null;
  const w = words(text); const res = parseRepeatAt(w, 0, anchor || startOfDay(new Date()));
  return res && res.used === w.length ? res.rule : null;
}
function startOfWeek(d) { return addDays(startOfDay(d), -((d.getDay() + 6) % 7)); }
function nthWeekday(monthStart, dow, ord) {
  const dim = daysInMonth(monthStart.getFullYear(), monthStart.getMonth()); const hits = [];
  for (let i = 0; i < dim; i++) { const d = addDays(monthStart, i); if (d.getDay() === dow) hits.push(d); }
  if (ord === -1) return hits[hits.length - 1];
  return ord <= hits.length ? hits[ord - 1] : null;
}
function lastWeekdayOf(monthStart) {
  let d = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0);
  while (d.getDay() === 0 || d.getDay() === 6) d = addDays(d, -1);
  return d;
}
/** The first date strictly after `after` that the rule lands on. */
function nextDate(r, after) {
  const d = startOfDay(after);
  switch (r.type) {
    case "days": {
      if (!r.weekdaysOnly) return addDays(d, r.interval);
      let x = addDays(d, 1); while (x.getDay() === 0 || x.getDay() === 6) x = addDays(x, 1); return x;
    }
    case "weeks": {
      const days = r.days.length ? r.days : [after.getDay()];
      for (let k = 1; k <= 7; k++) {
        const c = addDays(d, k); const same = +startOfWeek(c) === +startOfWeek(d);
        if (days.includes(c.getDay()) && (same || r.interval === 1)) return c;
        if (!same) break;
      }
      const ws = addDays(startOfWeek(d), 7 * r.interval);
      for (let k = 0; k < 7; k++) { const c = addDays(ws, k); if (days.includes(c.getDay())) return c; }
      return addDays(d, 7);
    }
    case "monthDay":
      for (let m = 0; m <= 24; m++) {
        if (m > 0 && m % r.interval !== 0) continue;
        const ms = addMonths(new Date(d.getFullYear(), d.getMonth(), 1), m);
        const dim = daysInMonth(ms.getFullYear(), ms.getMonth());
        const c = new Date(ms.getFullYear(), ms.getMonth(), r.monthDay === -1 ? dim : Math.min(r.monthDay, dim));
        if (c > d) return c;
      }
      return addMonths(d, r.interval);
    case "monthWeekday":
    case "monthLastWeekday":
      for (let m = 0; m <= 24; m++) {
        if (m > 0 && m % r.interval !== 0) continue;
        const ms = addMonths(new Date(d.getFullYear(), d.getMonth(), 1), m);
        const c = r.type === "monthLastWeekday" ? lastWeekdayOf(ms) : nthWeekday(ms, r.weekday, r.ordinal);
        if (c && c > d) return c;
      }
      return addMonths(d, r.interval);
    case "years": { const c = new Date(d); c.setFullYear(c.getFullYear() + r.interval); return c; }
  }
  return addDays(d, 1);
}
function describeRepeat(r) {
  const cap = s => s[0].toUpperCase() + s.slice(1);
  const every = r.interval === 1 ? "Every" : r.interval === 2 ? "Every other" : `Every ${r.interval}`;
  const plural = u => r.interval > 2 ? u + "s" : u;
  const sfx = n => [11, 12, 13].includes(n) ? "th" : ({ 1: "st", 2: "nd", 3: "rd" }[n % 10] || "th");
  const ord = o => o === -1 ? "last" : ORDINALS[o - 1];
  const joinDays = ds => {
    const n = [...ds].sort((a, b) => (a + 6) % 7 - (b + 6) % 7).map(i => cap(DAY_NAMES[i]));
    return n.length <= 1 ? n.join("") : n.slice(0, -1).join(", ") + " and " + n[n.length - 1];
  };
  switch (r.type) {
    case "days": return r.weekdaysOnly ? "Every weekday" : r.interval === 1 ? "Every day" : `${every} ${plural("day")}`;
    case "weeks": return (r.interval === 1 ? "Every " : `${every} ${plural("week")} on `) + joinDays(r.days);
    case "monthDay": return (r.interval === 1 ? "Every month" : `${every} ${plural("month")}`) + (r.monthDay === -1 ? " on the last day" : ` on the ${r.monthDay}${sfx(r.monthDay)}`);
    case "monthWeekday": return (r.interval === 1 ? "Every month" : `${every} ${plural("month")}`) + ` on the ${ord(r.ordinal)} ${cap(DAY_NAMES[r.weekday])}`;
    case "monthLastWeekday": return "Every month on the last weekday";
    case "years": return r.interval === 1 ? "Every year" : `${every} ${plural("year")}`;
  }
  return "Repeats";
}

// ---------- Reading what you type (same shortcuts as Focus) ----------

/** #tag, ! (high), !low, @today, due:fri, and everyday words: tomorrow, on friday, at 3pm, every friday. */
function applyInput(task, input, today) {
  today = startOfDay(today || new Date());
  const orig = input.split(/\s+/).filter(Boolean), low = orig.map(x => x.toLowerCase());
  const keep = []; let date = null, time = null, repeat = null, repeatText = "";
  task.Tags = task.Tags || [];
  for (let i = 0; i < orig.length; i++) {
    const wd = orig[i], lw = low[i];
    if (wd.length > 1 && wd[0] === "#") { task.Tags.push(wd.slice(1).replace(/[,.]+$/, "")); continue; }
    if (lw === "!" || lw === "!high" || lw === "!!") { task.Priority = "High"; continue; }
    if (lw === "!low") { task.Priority = "Low"; continue; }
    if (lw === "@today") { task.PlannedFor = ymd(today); continue; }
    if (lw.startsWith("due:")) { const r = tryDate(words(wd.slice(4)), 0, today, false); if (r) { date = r.date; continue; } }
    if (!repeat && (lw === "every" || lw === "each")) {
      const r = parseRepeatAt(low, i, date || today);
      if (r && r.used > 0) { repeat = r.rule; repeatText = low.slice(i, i + r.used).join(" "); i += r.used - 1; continue; }
    }
    if (!time) { const t = tryTime(low, i); if (t) { time = t; i += t.used - 1; continue; } }
    const skip = ["on", "by", "due"].includes(lw) ? 1 : 0;
    if (!date) { const r = tryDate(low, i + skip, today, true); if (r) { date = r.date; i += skip + r.used - 1; continue; } }
    keep.push(wd);
  }
  if (repeat) {
    task.Repeat = repeatText;
    if (!date) date = repeat.type === "days" && !repeat.weekdaysOnly ? today : nextDate(repeat, addDays(today, -1));
  }
  if (time) { if (!date) date = today; const r = new Date(date); r.setHours(time.h, time.m, 0, 0); task.RemindAt = iso(r); }
  if (date) task.Due = iso(startOfDay(date));
  task.Text = keep.join(" ");
  const seen = new Set();
  task.Tags = task.Tags.filter(t => t && !seen.has(t.toLowerCase()) && seen.add(t.toLowerCase()));
  return task;
}

// ---------- Task facts ----------

function dueOf(t) { const d = parseDate(t.Due); return d ? startOfDay(d) : null; }
function isOverdue(t, today) { const d = dueOf(t); return !!d && d < today; }
function isDueToday(t, today) { return sameDay(dueOf(t), today); }
function isPlannedToday(t, today) { return t.PlannedFor === ymd(today); }
function isSnoozed(t, today) { const s = parseDate(t.SnoozedUntil); return !!s && startOfDay(s) > today; }
function ageDays(t, today) { const c = parseDate(t.CreatedAt); return c ? Math.round((today - startOfDay(c)) / 864e5) : 0; }
/** Same order as Focus: overdue, due today, today's picks (high first), high priority, by due date. */
function rank(t, today) {
  const high = t.Priority === "High", lowp = t.Priority === "Low";
  return isOverdue(t, today) ? 0 : isDueToday(t, today) ? 1 : isPlannedToday(t, today) && high ? 2 : isPlannedToday(t, today) ? 3
       : high ? 4 : lowp ? 7 : dueOf(t) ? 5 : 6;
}
function sortTasks(list, today) {
  return [...list].sort((a, b) => rank(a, today) - rank(b, today) || ((dueOf(a) || 8.64e15) - (dueOf(b) || 8.64e15)));
}
function estimateLabel(e) { return !e ? "" : e < 1 ? "quick" : e === 1 ? "1 session" : `${Math.round(e)} sessions`; }
function details(t, today, projects) {
  const parts = [], fmt = (d, o) => d.toLocaleDateString([], o);
  const p = projects && projects.find(x => x.Id === t.Project); if (p) parts.push(p.Name);
  const d = dueOf(t);
  if (d && d < today) parts.push(`Overdue (was due ${fmt(d, { weekday: "short", month: "short", day: "numeric" })})`);
  else if (sameDay(d, today)) parts.push("Due today");
  else if (d) parts.push(d < addDays(today, 7) ? `Due ${fmt(d, { weekday: "long" })}` : `Due ${fmt(d, { month: "short", day: "numeric" })}`);
  if (isSnoozed(t, today)) parts.push(`Snoozed until ${fmt(parseDate(t.SnoozedUntil), { weekday: "short", month: "short", day: "numeric" })}`);
  const r = parseDate(t.RemindAt);
  if (r) parts.push(`Reminder ${sameDay(r, today) ? "" : fmt(r, { weekday: "short" }) + " "}${r.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`);
  if (t.Repeat) { const rule = parseRepeat(t.Repeat, d || today); parts.push(rule ? describeRepeat(rule) : "Repeats"); }
  if (t.Status === "doing") parts.push("Doing");
  if (t.Status === "waiting") parts.push(t.WaitingOn ? `Waiting on ${t.WaitingOn}` : "Waiting");
  if (t.Subtasks && t.Subtasks.length) parts.push(`${t.Subtasks.filter(s => s.Done).length} of ${t.Subtasks.length} steps`);
  if (t.Notes && t.Notes.trim()) parts.push("Has notes");
  if (t.Estimate > 0) parts.push(estimateLabel(t.Estimate));
  if (t.PlannedFor && !isPlannedToday(t, today)) {
    const pd = parseYmd(t.PlannedFor);
    if (pd && pd > today) parts.push(`Planned ${fmt(pd, { weekday: "short", month: "short", day: "numeric" })}`);
  }
  (t.Tags || []).forEach(x => parts.push("#" + x));
  return parts.join(", ");
}

/** Finishing a repeating task: the next one, with its new due date (and reminder), like Focus. */
function nextRepeat(t, today) {
  const rule = t.Repeat ? parseRepeat(t.Repeat, dueOf(t) || today) : null;
  if (!rule) return null;
  let next = nextDate(rule, dueOf(t) || today);
  while (next < today) next = nextDate(rule, next);
  const again = JSON.parse(JSON.stringify(t));
  again.Id = newId(); again.Due = iso(next); again.From = "Repeats: " + describeRepeat(rule);
  const r = parseDate(t.RemindAt);
  again.RemindAt = r ? iso(new Date(next.getFullYear(), next.getMonth(), next.getDate(), r.getHours(), r.getMinutes())) : null;
  again.Subtasks = (t.Subtasks || []).map(s => ({ Text: s.Text, Done: false }));
  again.PlannedFor = ""; again.SnoozedUntil = null; again.Status = ""; again.Spent = 0; again.CreatedAt = iso(new Date());
  return again;
}
function newId() {
  const s = (globalThis.crypto && crypto.randomUUID) ? crypto.randomUUID()
          : (Date.now().toString(16) + Math.random().toString(16).slice(2) + "00000000000000000000000000000000");
  return s.replace(/-/g, "").slice(0, 32);
}
if (typeof module !== "undefined") module.exports = { applyInput, parseRepeat, nextDate, describeRepeat, tryDate, tryTime, details, rank, nextRepeat, iso, ymd, startOfDay, addDays };
