// Focus Tasks: the app (screens, signing in, reading and saving the task files).
"use strict";
// Settings live in config.js.
const CLIENT_ID = String((window.FOCUS_CONFIG || {}).clientId || "PASTE");
const FOLDER = String((window.FOCUS_CONFIG || {}).folder || "Focus");

// Hardening: the page may only ever touch these three files, in the Focus folder, and nothing else.
const ALLOWED_FILES = new Set(["task-bin.json", "resolved-tasks.json", "projects.json"]);
// Read only: the calendar Focus shares, and day plans (for the calendar view).
const READ_ONLY = name => name === "calendar-share.json" || /^plans\/\d{4}-\d{2}-\d{2}\.json$/.test(name);
function fileUrl(name, writing) {
  if (!ALLOWED_FILES.has(name) && (writing || !READ_ONLY(name))) throw new Error("Blocked: this page only works with Focus's task files.");
  if (!/^[A-Za-z0-9 _-]{1,64}$/.test(FOLDER)) throw new Error("Blocked: the folder name in config.js isn't allowed.");
  return GRAPH + encodeURIComponent(FOLDER) + "/" + name.split("/").map(encodeURIComponent).join("/") + ":/content";
}

const SCOPES = ["Files.ReadWrite"];
const GRAPH = "https://graph.microsoft.com/v1.0/me/drive/root:/";
const $ = id => document.getElementById(id);

let client, account;
let bin = [], resolved = [], projects = [];
let view = "today", chip = "all", search = "", busy = false;

const today = () => startOfDay(new Date());
function setStatus(text, isError) { $("status").textContent = text || ""; $("status").classList.toggle("error", !!isError); }
function toast(text) {
  const t = document.createElement("div"); t.className = "toast"; t.textContent = text; document.body.append(t);
  setTimeout(() => t.remove(), 2200);
}
function esc(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

// ---------- Signing in ----------

async function start() {
  if (CLIENT_ID.startsWith("PASTE")) { setStatus("This page isn't set up yet: add your client ID (see SETUP.md).", true); return; }
  client = new window.msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: "https://login.microsoftonline.com/common", redirectUri: location.origin + location.pathname },
    // Remembered on this phone, so you stay signed in between visits. Sign out with ⎋ to forget it.
    cache: { cacheLocation: "localStorage" }
  });
  try { const r = await client.handleRedirectPromise(); account = (r && r.account) || client.getAllAccounts()[0]; }
  catch (e) { setStatus("Sign-in problem: " + e.message, true); }
  showSignedIn(!!account);
  if (account) await reload();
}
function showSignedIn(on) {
  $("signedOut").classList.toggle("hidden", on);
  $("app").classList.toggle("hidden", !on);
  $("addbar").classList.toggle("hidden", !on);
  ["menuBtn", "projectsBtn", "calBtn"].forEach(id => $(id).classList.toggle("hidden", !on));
}
async function token() {
  try { return (await client.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
  catch {
    // Every so often Microsoft wants a quick check: this passes through its page and straight back,
    // normally without typing anything.
    setStatus("Refreshing your sign-in...");
    await client.acquireTokenRedirect({ scopes: SCOPES, account, loginHint: account && account.username });
    throw new Error("Refreshing your sign-in...");
  }
}

// ---------- The Focus files in OneDrive ----------

async function readFile(name) {
  const res = await fetch(fileUrl(name), { headers: { Authorization: "Bearer " + await token() }, cache: "no-store" });
  if (res.status === 404) return [];
  if (!res.ok) throw new Error(`Couldn't read ${name} (${res.status})`);
  const text = await res.text();
  try { const data = JSON.parse(text); return Array.isArray(data) ? data : (data && typeof data === "object" ? data : []); }
  catch { throw new Error(`${name} is busy syncing; try again in a moment.`); }
}
async function writeFile(name, data) {
  const res = await fetch(fileUrl(name, true), {
    method: "PUT", headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json" },
    body: JSON.stringify(data, null, 2)
  });
  if (!res.ok) throw new Error(`Couldn't save ${name} (${res.status})`);
}
async function reload() {
  try {
    setStatus("Loading...");
    [bin, resolved, projects] = await Promise.all([readFile("task-bin.json"), readFile("resolved-tasks.json"), readFile("projects.json")]);
    render();
    setStatus("Up to date, " + new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }));
  } catch (e) { setStatus(e.message, true); }
}
/** Every change is made on the newest copy of the files, so edits from your computers aren't lost. */
async function change(fn, message) {
  if (busy) { toast("One moment..."); return; }
  busy = true;
  try {
    setStatus("Saving...");
    const [b, r, p] = await Promise.all([readFile("task-bin.json"), readFile("resolved-tasks.json"), readFile("projects.json")]);
    const touched = fn(b, r, p) || {};
    for (const t of b) if (!t.CreatedAt) t.CreatedAt = iso(new Date());
    await writeFile("task-bin.json", b);
    if (touched.resolved) await writeFile("resolved-tasks.json", r);
    if (touched.projects) await writeFile("projects.json", p);
    bin = b; resolved = r; projects = p;
    render();
    setStatus("Saved. Your computers pick it up within a minute or so.");
    if (message) toast(message);
  } catch (e) { setStatus(e.message, true); }
  finally { busy = false; }
}

// ---------- Lists ----------

function visible(t) {
  const d = today();
  if (search) {
    const q = search.toLowerCase();
    const hit = (t.Text || "").toLowerCase().includes(q) || (t.Notes || "").toLowerCase().includes(q) ||
      (t.Tags || []).some(x => x.toLowerCase().includes(q.replace(/^#/, ""))) ||
      (t.Subtasks || []).some(s => (s.Text || "").toLowerCase().includes(q));
    if (!hit) return false;
  }
  if (chip === "overdue" && !isOverdue(t, d)) return false;
  if (chip === "high" && t.Priority !== "High") return false;
  if (chip === "age7" && ageDays(t, d) < 7) return false;
  if (chip.startsWith("#") && !(t.Tags || []).some(x => x.toLowerCase() === chip.slice(1).toLowerCase())) return false;
  return true;
}
function projectById(id) { return projects.find(p => p.Id === id); }
function isToday(t) { const d = today(); return isPlannedToday(t, d) || isDueToday(t, d) || isOverdue(t, d); }

function render() {
  const d = today();
  const monday = addDays(d, -((d.getDay() + 6) % 7));
  const doneToday = resolved.filter(r => (parseDate(r.ResolvedAt) || 0) >= d).length;
  const doneWeek = resolved.filter(r => (parseDate(r.ResolvedAt) || 0) >= monday).length;
  $("stats").innerHTML = `<b>${doneToday}</b> done today, <b>${doneWeek}</b> this week, <b>${bin.length}</b> to do`;
  document.querySelectorAll("#tabs button").forEach(b => b.setAttribute("aria-pressed", String(b.dataset.view === view)));
  renderChips();
  const list = $("list"); list.innerHTML = "";

  $("projectBar").classList.toggle("hidden", !view.startsWith("project:"));
  if (view === "done") return renderDone(list);
  const tasks = bin.filter(t => view === "upcoming" || !isSnoozed(t, d)).filter(visible);
  if (view === "upcoming") return renderUpcoming(list, tasks);
  if (view.startsWith("project:")) return renderProject(list, tasks, view.slice(8));
  const shown = sortTasks(view === "today" ? tasks.filter(isToday) : view === "undated" ? tasks.filter(isUndated) : tasks, d);
  if (!shown.length) {
    list.innerHTML = `<p class="empty">${view === "today"
      ? "Nothing set for today. Add a task below, or open one in All and choose Do it: Today."
      : view === "undated" ? "Every task has a date. Nice."
      : "No tasks here."}</p>`;
    return;
  }
  const ul = document.createElement("ul"); shown.forEach(t => ul.append(row(t))); list.append(ul);
}
/** One project's tasks on their own, grouped by section (like a Todoist project). */
function renderProject(list, tasks, id) {
  const d = today(), p = projectById(id);
  if (!p) { view = "all"; return render(); }
  $("projectTitle").innerHTML = `<span class="dot" style="background:${esc(p.Color)}"></span>${esc(p.Name)}`;
  const mine = sortTasks(tasks.filter(t => t.Project === id), d);
  if (!mine.length) { list.innerHTML = `<p class="empty">No tasks in ${esc(p.Name)}. Add one below.</p>`; return; }
  const secs = p.Sections || [];
  const groups = [["", "No section"], ...secs.map(x => [x.Id, x.Name])];
  for (const [sid, name] of groups) {
    const items = mine.filter(t => sid === "" ? !secs.some(x => x.Id === t.Section) : t.Section === sid);
    if (!items.length) continue;
    if (secs.length) { const h = document.createElement("h2"); h.textContent = name; list.append(h); }
    const ul = document.createElement("ul"); items.forEach(t => ul.append(row(t))); list.append(ul);
  }
}
function openProjects() {
  const counts = id => bin.filter(t => t.Project === id).length;
  const card = $("sheetCard");
  card.innerHTML = `<h3 style="margin:4px 0 10px">Projects</h3>` + (projects.length
    ? `<ul>${projects.map(p => `<li class="menurow" data-id="${esc(p.Id)}"><span class="dot" style="background:${esc(p.Color)}"></span>${esc(p.Name)}<span class="count">${counts(p.Id)}</span></li>`).join("")}</ul>`
    : `<p class="empty">No projects yet. Type #Name when adding a task to create one.</p>`) +
    `<div class="sheetbar"><span></span><button type="button" id="pClose">Close</button></div>`;
  card.querySelectorAll(".menurow").forEach(li => li.onclick = () => { view = "project:" + li.dataset.id; chip = "all"; closeSheet(); render(); });
  $("pClose").onclick = closeSheet;
  openSheet();
}
function renderChips() {
  const d = today(), chips = [["all", "All"], ["overdue", `Overdue (${bin.filter(t => isOverdue(t, d)).length})`],
    ["high", "High priority"], ["age7", "+ one week old"]];
  const tags = [...new Set(bin.flatMap(t => t.Tags || []).map(x => x.toLowerCase()))].sort();
  tags.forEach(x => chips.push(["#" + x, "#" + esc(x)]));
  if (!chips.some(c => c[0] === chip)) chip = "all";
  $("chips").classList.toggle("hidden", view === "done");
  $("chips").innerHTML = chips.map(([k, l]) => `<button data-chip="${esc(k)}" aria-pressed="${k === chip}">${l}</button>`).join("");
  $("chips").querySelectorAll("button").forEach(b => b.onclick = () => { chip = b.dataset.chip; render(); });
}
function row(t) {
  const d = today(), li = document.createElement("li"); li.className = "task";
  const check = document.createElement("button");
  check.className = "check" + (t.Priority === "High" ? " high" : t.Priority === "Low" ? " low" : "");
  check.textContent = "✓"; check.setAttribute("aria-label", "Finish");
  check.onclick = () => { check.classList.add("done"); finish(t.Id); };
  const body = document.createElement("div"); body.className = "body";
  const det = details(t, d, projects);
  body.innerHTML = `<div class="title">${esc(t.Text)}</div>` + (det ? `<div class="details${isOverdue(t, d) ? " overdue" : ""}">${esc(det)}</div>` : "");
  body.onclick = () => openEditor(t.Id);
  li.append(check, body); return li;
}
function renderUpcoming(list, tasks) {
  const d = today(), groups = new Map();
  const dateOf = t => {
    if (isSnoozed(t, d)) return startOfDay(parseDate(t.SnoozedUntil));
    const p = parseYmd(t.PlannedFor); if (p && p > d) return p;
    return dueOf(t);
  };
  const add = (h, t) => { if (!groups.has(h)) groups.set(h, []); groups.get(h).push(t); };
  sortTasks(tasks, d).map(t => [t, dateOf(t)]).filter(x => x[1]).sort((a, b) => a[1] - b[1]).forEach(([t, when]) => {
    if (when < d) add("Overdue", t);
    else if (when < addDays(d, 14)) add(sameDay(when, d) ? "Today" : sameDay(when, addDays(d, 1)) ? "Tomorrow"
      : when.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" }), t);
    else add("Later", t);
  });
  if (!groups.size) { list.innerHTML = `<p class="empty">Nothing scheduled. Give tasks a due date or snooze them to see them here.</p>`; return; }
  const order = ["Overdue", ...[...groups.keys()].filter(k => k !== "Overdue" && k !== "Later"), "Later"];
  for (const k of order) if (groups.has(k)) {
    const h = document.createElement("h2"); h.textContent = k; list.append(h);
    const ul = document.createElement("ul"); groups.get(k).forEach(t => ul.append(row(t))); list.append(ul);
  }
}
function renderDone(list) {
  const d = today(), groups = new Map();
  resolved.filter(r => !search || (r.Text || "").toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => (parseDate(b.ResolvedAt) || 0) - (parseDate(a.ResolvedAt) || 0)).slice(0, 300).forEach(r => {
      const day = startOfDay(parseDate(r.ResolvedAt) || 0);
      const h = sameDay(day, d) ? "Today" : sameDay(day, addDays(d, -1)) ? "Yesterday"
        : day.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" });
      if (!groups.has(h)) groups.set(h, []); groups.get(h).push(r);
    });
  if (!groups.size) { list.innerHTML = `<p class="empty">Nothing finished yet.</p>`; return; }
  for (const [h, items] of groups) {
    const head = document.createElement("h2"); head.textContent = h; list.append(head);
    const ul = document.createElement("ul");
    items.forEach(r => {
      const li = document.createElement("li"); li.className = "task";
      li.innerHTML = `<span class="check done">✓</span><div class="body"><div class="title" style="text-decoration:line-through;color:var(--muted)">${esc(r.Text)}</div></div>`;
      const back = document.createElement("button"); back.className = "link"; back.textContent = "Reopen";
      back.onclick = () => reopen(r.Id); li.append(back); ul.append(li);
    });
    list.append(ul);
  }
}

// ---------- Actions ----------

function add(input) {
  const d = today();
  const { task, created, notice } = applyInput({
    Id: newId(), Text: "", From: "Added on phone, " + d.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" }),
    Priority: "Normal", Tags: [], Subtasks: [], Notes: "", Repeat: "", Status: "", WaitingOn: "", Project: "", PlannedFor: "",
    Estimate: 0, Spent: 0, CreatedAt: iso(new Date())
  }, input, d, projects);
  if (!task.Text) return;
  if (view === "today" && !task.PlannedFor && !task.Due) task.PlannedFor = ymd(d); // added on Today: it's for today
  if (view.startsWith("project:") && !task.Project) task.Project = view.slice(8);
  change((b, r, p) => {
    b.push(task);
    if (created && !p.some(q => q.Id === created.Id)) { p.push(created); return { projects: true }; }
    return {};
  }, notice || (created ? `Added, in a new project: ${created.Name}` : "Added"));
}
function finish(id) {
  change((b, r) => {
    const i = b.findIndex(x => x.Id === id); if (i < 0) return {};
    const [t] = b.splice(i, 1);
    r.push({ Id: newId(), Text: t.Text, From: t.From || "", ResolvedAt: iso(new Date()), Spent: t.Spent || 0, Estimate: t.Estimate || 0, Project: t.Project || "" });
    const again = nextRepeat(t, today()); if (again) b.push(again);
    return { resolved: true };
  }, "Done ✓");
}
function reopen(id) {
  change((b, r) => {
    const i = r.findIndex(x => x.Id === id); if (i < 0) return {};
    const [t] = r.splice(i, 1);
    b.push({ Id: newId(), Text: t.Text, From: "Reopened on phone", Priority: "Normal", Tags: [], Subtasks: [], Project: t.Project || "", CreatedAt: iso(new Date()) });
    return { resolved: true };
  }, "Back in your list");
}

// ---------- Editing a task ----------

function openEditor(id) {
  const t = bin.find(x => x.Id === id); if (!t) return;
  const d = today(), due = dueOf(t), remind = parseDate(t.RemindAt), snooze = parseDate(t.SnoozedUntil);
  const planned = t.PlannedFor || "", tdy = ymd(d), tmr = ymd(addDays(d, 1));
  const seg = (name, options, value) => `<div class="seg" data-seg="${name}">` +
    options.map(([v, l]) => `<button type="button" data-v="${esc(v)}" aria-pressed="${String(v) === String(value)}">${l}</button>`).join("") + "</div>";
  const steps = JSON.parse(JSON.stringify(t.Subtasks || []));
  const planChoice = planned === "" ? "" : planned === tdy || planned === tmr ? planned : "other";
  const card = $("sheetCard");
  card.innerHTML = `
    <label>Task</label><textarea id="eText" rows="2">${esc(t.Text)}</textarea>
    <label>Do it</label>${seg("plan", [["", "Any time"], [tdy, "Today"], [tmr, "Tomorrow"], ["other", "Pick a day"]], planChoice)}
    <input id="ePlanDay" type="date" class="${planChoice === "other" ? "" : "hidden"}" value="${esc(planChoice === "other" ? planned : "")}" style="margin-top:6px">
    <label>Priority</label>${seg("prio", [["High", "High"], ["Normal", "Normal"], ["Low", "Low"]], t.Priority || "Normal")}
    <div class="row">
      <div><label>Due date</label><input id="eDue" type="date" value="${due ? ymd(due) : ""}"></div>
      <div><label>Reminder</label><input id="eRemind" type="datetime-local" value="${remind ? iso(remind).slice(0, 16) : ""}"></div>
    </div>
    <label>Repeats</label><input id="eRepeat" value="${esc(t.Repeat || "")}" placeholder="every friday, weekdays, every month on the 15th">
    <div class="hint" id="eRepeatHint"></div>
    <label>Project</label><select id="eProject"><option value="">None</option>${projects.map(p =>
      `<option value="${esc(p.Id)}" ${p.Id === t.Project ? "selected" : ""}>${esc(p.Name)}</option>`).join("")}<option value="__new">+ New project...</option></select>
    <label>Section</label><select id="eSection"></select>
    <label>Tags</label><input id="eTags" value="${esc((t.Tags || []).join(", "))}" placeholder="client, admin">
    <label>Steps</label><ul class="steps" id="eSteps"></ul>
    <input id="eStepNew" placeholder="Add a step and press Enter" enterkeyhint="done">
    <label>Notes</label><textarea id="eNotes">${esc(t.Notes || "")}</textarea>
    <label>Estimate (focus sessions)</label>${seg("est", [[0, "Not sure"], [0.34, "Quick"], [1, "1"], [2, "2"], [3, "3"], [4, "4"]], t.Estimate || 0)}
    <label>Status</label>${seg("status", [["", "To do"], ["doing", "Doing"], ["waiting", "Waiting"]], t.Status || "")}
    <input id="eWaiting" class="${t.Status === "waiting" ? "" : "hidden"}" value="${esc(t.WaitingOn || "")}" placeholder="Waiting on whom?" style="margin-top:6px">
    <label>Snooze until (hidden from lists until then)</label><input id="eSnooze" type="date" value="${snooze && snooze > d ? ymd(snooze) : ""}">
    <div class="sheetbar">
      <button type="button" id="eDelete">Delete</button>
      <span><button type="button" id="eCancel">Cancel</button> <button type="button" class="primary" id="eSave">Save</button></span>
    </div>`;
  openSheet();

  card.querySelectorAll(".seg").forEach(s => s.querySelectorAll("button").forEach(b => b.onclick = () => {
    s.querySelectorAll("button").forEach(x => x.setAttribute("aria-pressed", "false"));
    b.setAttribute("aria-pressed", "true");
    if (s.dataset.seg === "plan") $("ePlanDay").classList.toggle("hidden", b.dataset.v !== "other");
    if (s.dataset.seg === "status") $("eWaiting").classList.toggle("hidden", b.dataset.v !== "waiting");
  }));
  const segValue = name => { const b = card.querySelector(`[data-seg="${name}"] [aria-pressed="true"]`); return b ? b.dataset.v : ""; };
  const repeatHint = () => {
    const txt = $("eRepeat").value.trim(), rule = parseRepeat(txt, $("eDue").value ? parseYmd($("eDue").value) : d);
    $("eRepeatHint").textContent = !txt ? "Doesn't repeat" : rule ? describeRepeat(rule) : "Not understood yet";
  };
  $("eRepeat").oninput = repeatHint; repeatHint();
  const drawSteps = () => {
    $("eSteps").innerHTML = "";
    steps.forEach((s, i) => {
      const li = document.createElement("li");
      li.innerHTML = `<input type="checkbox" ${s.Done ? "checked" : ""} style="width:auto"><span>${esc(s.Text)}</span><button type="button" class="icon">✕</button>`;
      li.querySelector("input").onchange = e => { s.Done = e.target.checked; };
      li.querySelector("button").onclick = () => { steps.splice(i, 1); drawSteps(); };
      $("eSteps").append(li);
    });
  };
  drawSteps();
  $("eStepNew").onkeydown = e => {
    if (e.key !== "Enter") return;
    e.preventDefault(); const v = e.target.value.trim();
    if (v) { steps.push({ Text: v, Done: false }); e.target.value = ""; drawSteps(); }
  };
  let newProject = null;
  $("eProject").onchange = e => {
    if (e.target.value !== "__new") return;
    const name = (prompt("New project name") || "").trim();
    if (!name) { e.target.value = t.Project || ""; return; }
    const pal = ["#DC4C3E", "#246FE0", "#058527", "#EB8909", "#884DFF", "#14AAF5", "#B8256F", "#6ACCBC"];
    newProject = { Id: newId(), Name: name, Color: pal[projects.length % pal.length], Aliases: [], Sections: [] };
    const opt = document.createElement("option"); opt.value = newProject.Id; opt.textContent = name;
    e.target.insertBefore(opt, e.target.lastElementChild); e.target.value = newProject.Id;
  };
  const fillSections = () => {
    const id = $("eProject").value;
    const p = projectById(id) || (newProject && newProject.Id === id ? newProject : null);
    const secs = (p && p.Sections) || [];
    $("eSection").innerHTML = `<option value="">None</option>` + secs.map(x => `<option value="${esc(x.Id)}" ${x.Id === t.Section ? "selected" : ""}>${esc(x.Name)}</option>`).join("");
    $("eSection").disabled = !secs.length;
  };
  fillSections();
  $("eProject").addEventListener("change", fillSections);
  const close = closeSheet;
  $("eCancel").onclick = close;
  $("eDelete").onclick = () => {
    if (!confirm("Delete this task?")) return;
    close(); change(b => { const i = b.findIndex(x => x.Id === id); if (i >= 0) b.splice(i, 1); }, "Deleted");
  };
  $("eSave").onclick = () => {
    const text = $("eText").value.trim(); if (!text) { toast("The task needs some text"); return; }
    const repeat = $("eRepeat").value.trim();
    if (repeat && !parseRepeat(repeat, d)) { toast("Couldn't read the repeat"); return; }
    let plan = segValue("plan"); if (plan === "other") plan = $("ePlanDay").value || "";
    const status = segValue("status");
    const edits = {
      Text: text, PlannedFor: plan, Priority: segValue("prio") || "Normal",
      Due: $("eDue").value ? iso(parseYmd($("eDue").value)) : null,
      RemindAt: $("eRemind").value ? $("eRemind").value + ":00" : null,
      Repeat: repeat, Project: $("eProject").value === "__new" ? "" : $("eProject").value,
      Section: $("eSection").disabled ? "" : $("eSection").value,
      Tags: $("eTags").value.split(",").map(x => x.trim().replace(/^#/, "")).filter(Boolean),
      Subtasks: steps, Notes: $("eNotes").value, Estimate: +segValue("est") || 0,
      Status: status, WaitingOn: status === "waiting" ? $("eWaiting").value.trim() : "",
      SnoozedUntil: $("eSnooze").value ? iso(parseYmd($("eSnooze").value)) : null
    };
    if (repeat && !edits.Due) edits.Due = iso(nextDate(parseRepeat(repeat, d), addDays(d, -1)));
    if (status === "waiting" && t.Status !== "waiting") edits.WaitingSince = iso(new Date());
    const projectToAdd = newProject && edits.Project === newProject.Id ? newProject : null;
    close();
    change((b, r, p) => {
      const x = b.find(y => y.Id === id); if (!x) return {};
      Object.assign(x, edits);
      if (projectToAdd && !p.some(q => q.Id === projectToAdd.Id)) { p.push(projectToAdd); return { projects: true }; }
      return {};
    }, "Saved");
  };
}

// ---------- Sheets ----------

function openSheet() { $("sheet").classList.remove("hidden"); }
function closeSheet() { $("sheet").classList.add("hidden"); }

// ---------- Calendar (kept small: tasks stay the main thing) ----------

let calDay = null;
const calShowPlan = () => localStorage.getItem("focus-cal-plan") !== "off";

/** Meetings (shared by Focus on your computer) and, if you like, that day's plan. */
async function openCalendar(day) {
  calDay = startOfDay(day || calDay || new Date());
  const card = $("sheetCard");
  card.innerHTML = `<div class="calhead"><button type="button" class="icon" id="cPrev" aria-label="Previous day">‹</button>
      <b id="cTitle"></b><button type="button" class="icon" id="cNext" aria-label="Next day">›</button></div>
    <div class="calbar"><button type="button" id="cToday">Today</button>
      <label class="switch"><input type="checkbox" id="cPlan" ${calShowPlan() ? "checked" : ""}> Show my plan</label></div>
    <div id="cList" class="callist"><p class="empty">Loading...</p></div>
    <div class="hint" id="cNote"></div>
    <div class="sheetbar"><span></span><button type="button" id="cClose">Close</button></div>`;
  $("cTitle").textContent = sameDay(calDay, today()) ? "Today" : calDay.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" });
  $("cPrev").onclick = () => openCalendar(addDays(calDay, -1));
  $("cNext").onclick = () => openCalendar(addDays(calDay, 1));
  $("cToday").onclick = () => openCalendar(new Date());
  $("cPlan").onchange = e => { localStorage.setItem("focus-cal-plan", e.target.checked ? "on" : "off"); openCalendar(calDay); };
  $("cClose").onclick = closeSheet;
  openSheet();
  const shownDay = calDay;
  try {
    const shared = await readFile("calendar-share.json");
    if (shownDay !== calDay) return; // you've moved to another day meanwhile
    const items = (Array.isArray(shared) ? shared : [])
      .map(e => ({ title: e.Title || "Busy", start: parseDate(e.Start), end: parseDate(e.End), allDay: !!e.AllDay, busy: e.Busy !== false, kind: "meeting" }))
      .filter(e => e.start && e.end && e.start < addDays(shownDay, 1) && e.end > shownDay);
    if (calShowPlan()) {
      const plan = await readFile(`plans/${ymd(shownDay)}.json`);
      if (plan && !Array.isArray(plan) && Array.isArray(plan.Blocks)) {
        const [h, m] = String(plan.DayStart || "06:45").split(":").map(Number);
        let t = new Date(shownDay); t.setHours(h || 0, m || 0, 0, 0);
        for (const b of plan.Blocks) {
          const start = new Date(t), end = new Date(t.getTime() + (b.Minutes || 0) * 60000); t = end;
          if (b.Kind === "Meeting") continue; // meetings come from the calendar itself
          items.push({ title: b.Name || (b.Kind === "Focus" ? "Focus" : "Break"), start, end, kind: b.Kind === "Focus" ? "focus" : "break",
                       tasks: (b.Tasks || []).filter(x => x && x.Text).map(x => x.Text) });
        }
      }
    }
    if (shownDay !== calDay) return;
    items.sort((a, b) => (a.allDay ? -1 : 0) - (b.allDay ? -1 : 0) || a.start - b.start || (a.kind === "meeting" ? -1 : 1));
    const now = new Date(), time = d => d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    let html = "", nowShown = !sameDay(shownDay, now);
    for (const it of items) {
      if (!nowShown && !it.allDay && it.start > now) { html += `<div class="nowline"><span>Now ${time(now)}</span></div>`; nowShown = true; }
      html += `<div class="cal ${it.kind}${it.busy === false ? " free" : ""}"><div class="when">${it.allDay ? "All day" : `${time(it.start)}<br><small>${time(it.end)}</small>`}</div>` +
        `<div class="what">${esc(it.title)}${it.tasks && it.tasks.length ? `<small>${it.tasks.map((x, i) => `${i + 1}. ${esc(x)}`).join("<br>")}</small>` : ""}</div></div>`;
    }
    if (!nowShown) html += `<div class="nowline"><span>Now ${time(now)}</span></div>`;
    $("cList").innerHTML = items.length ? html : `<p class="empty">Nothing on the calendar${calShowPlan() ? " or planned" : ""}.</p>`;
    $("cNote").textContent = Array.isArray(shared) && shared.length
      ? "Meetings come from Focus on your computer (updated every 15 minutes while it's on)."
      : "No meetings shared yet. Focus on your computer shares your calendar once it's connected and syncing.";
  } catch (e) { if (shownDay === calDay) $("cList").innerHTML = `<p class="empty">${esc(e.message)}</p>`; }
}

// ---------- Menu and themes ----------

const THEMES = [["classic", "Classic"], ["mono", "Monochrome"], ["calm", "Calm"]];
function applyTheme() {
  const k = localStorage.getItem("focus-theme") || "classic";
  document.documentElement.dataset.theme = THEMES.some(t => t[0] === k) ? k : "classic";
}
function openMenu() {
  const current = document.documentElement.dataset.theme || "classic";
  const card = $("sheetCard");
  card.innerHTML = `<label>Theme</label><div class="seg">${THEMES.map(([k, n]) => `<button type="button" data-k="${k}" aria-pressed="${k === current}">${n}</button>`).join("")}</div>
    <div class="hint">Light or dark follows your phone.</div>
    <div class="sheetbar"><button type="button" id="mSignout">Sign out</button><button type="button" id="mClose">Close</button></div>`;
  card.querySelectorAll(".seg button").forEach(b => b.onclick = () => { localStorage.setItem("focus-theme", b.dataset.k); applyTheme(); openMenu(); });
  $("mSignout").onclick = () => { closeSheet(); client.logoutRedirect({ account }); };
  $("mClose").onclick = closeSheet;
  openSheet();
}
applyTheme();

// ---------- Wiring up ----------

$("signin").onclick = () => client.loginRedirect({ scopes: SCOPES });
$("menuBtn").onclick = openMenu;
$("projectsBtn").onclick = openProjects;
$("calBtn").onclick = () => openCalendar(new Date());
$("projectBack").onclick = () => { view = "all"; render(); };
$("sheet").addEventListener("click", e => { if (e.target.id === "sheet") closeSheet(); });
$("refresh").onclick = reload;
$("searchBtn").onclick = () => {
  const box = $("search"); box.classList.toggle("hidden");
  if (!box.classList.contains("hidden")) box.focus(); else { search = ""; box.value = ""; render(); }
};
$("search").oninput = e => { search = e.target.value.trim(); render(); };
$("tabs").querySelectorAll("button").forEach(b => b.onclick = () => { view = b.dataset.view; render(); });
$("addForm").onsubmit = e => { e.preventDefault(); const v = $("newTask").value.trim(); if (!v) return; $("newTask").value = ""; add(v); };
// Get the latest whenever you come back to the page (unless you're in the middle of editing).
document.addEventListener("visibilitychange", () => { if (!document.hidden && account && $("sheet").classList.contains("hidden")) reload(); });

// Lets your phone install the page as an app (and open it offline to the last copy).
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => { /* still works as a normal page */ });

start();
