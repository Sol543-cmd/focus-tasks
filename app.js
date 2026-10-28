// Focus Tasks: the app (screens, signing in, reading and saving the task files).
"use strict";
// Settings live in config.js.
const CLIENT_ID = String((window.FOCUS_CONFIG || {}).clientId || "PASTE");
const FOLDER = String((window.FOCUS_CONFIG || {}).folder || "Focus");

// Hardening: the page may only ever touch these three files, in the Focus folder, and nothing else.
const ALLOWED_FILES = new Set(["task-bin.json", "resolved-tasks.json", "projects.json", "task-removed.json",
                               "personal-tasks.json", "personal-finished.json", "personal-projects.json", "personal-removed.json"]);
/** Two pages: your tasks, and (swipe) your personal tasks, each with its own three files. */
const FILESETS = {
  main: ["task-bin.json", "resolved-tasks.json", "projects.json", "task-removed.json"],
  personal: ["personal-tasks.json", "personal-finished.json", "personal-projects.json", "personal-removed.json"]
};
let mode = "personal"; // Focus opens on Personal; swipe left for Tasks
const files = () => FILESETS[mode];
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
/** What a page opens on: All, unless you chose Today in the ⋯ menu ("personal" or "main"). */
const openingView = m => localStorage.getItem("focus-open-" + m) === "today" ? "today" : "all";
let view = openingView("personal"), chip = "all", search = "", busy = false;
let filtersOpen = false;

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
  $("boot").classList.add("hidden"); // the app has started
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
    const wanted = mode;
    const loaded = await Promise.all(files().slice(0, 3).map(readFile));
    if (wanted !== mode) return; // you switched pages while it loaded
    [bin, resolved, projects] = loaded;
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
    const [binFile, doneFile, projectFile] = files();
    const [b, r, p] = await Promise.all([readFile(binFile), readFile(doneFile), readFile(projectFile)]);
    const before = new Map(b.map(t => [t.Id, JSON.stringify(t)]));
    const touched = fn(b, r, p) || {};
    // Every task added or changed here gets the time of the change (in the same form as your computers use),
    // so your computers always know it's newer than any old copy, and never mistake it for one.
    const stamp = new Date().toISOString();
    for (const t of b) {
      if (!t.CreatedAt) t.CreatedAt = iso(new Date());
      if (before.get(t.Id) !== JSON.stringify(t)) t.Modified = stamp;
    }
    // Tasks finished or deleted here go in the removal record, so your computers know they're really gone.
    const gone = [...before.keys()].filter(id => !b.some(t => t.Id === id));
    if (gone.length) {
      const removedFile = files()[3];
      const record = await readFile(removedFile);
      for (const id of gone) if (!record.some(x => x.Id === id)) record.push({ Id: id, At: stamp });
      await writeFile(removedFile, record);
    }
    await writeFile(binFile, b);
    if (touched.resolved) await writeFile(doneFile, r);
    if (touched.projects) await writeFile(projectFile, p);
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
  document.body.classList.toggle("filters-open", filtersOpen);
  $("filterDot").classList.toggle("hidden", chip === "all");
  $("calDate").textContent = new Date().toLocaleDateString([], { weekday: "short", day: "numeric" });
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

const calShowPlan = () => localStorage.getItem("focus-cal-plan") !== "off";

/** Meetings (shared by Focus on your computer) and, if you like, that day's plan. */
// ---------- Google Calendar (Personal's calendar), two-way ----------
// Connects straight to Google (no outside code is loaded): you sign in on Google's page, Google sends you back with
// a pass that lasts about an hour, and the app reads and changes your events with it. When it runs out, Connect again.

const GOOGLE_SCOPE = "https://www.googleapis.com/auth/calendar.events";
const gClientId = () => (localStorage.getItem("focus-google-client") || (window.FOCUS_CONFIG && window.FOCUS_CONFIG.googleClientId) || "").trim();
const gCalendarId = () => (localStorage.getItem("focus-google-calendar") || "primary").trim() || "primary";
function gToken() {
  try { const t = JSON.parse(localStorage.getItem("focus-google-token") || "null"); return t && t.exp > Date.now() + 60000 ? t.token : null; }
  catch { return null; }
}
function gConnect() {
  const id = gClientId();
  if (!id) { toast("Add your Google client ID first (⋯ menu, Google Calendar)."); return; }
  const back = location.origin + location.pathname;
  location.href = "https://accounts.google.com/o/oauth2/v2/auth?" + new URLSearchParams({
    client_id: id, redirect_uri: back, response_type: "token", scope: GOOGLE_SCOPE, include_granted_scopes: "true", state: "focus-google", prompt: ""
  }).toString().replace("&prompt=", "");
}
function gDisconnect() { localStorage.removeItem("focus-google-token"); toast("Google Calendar disconnected"); }
// Coming back from Google's sign-in page: keep the pass, and tidy the address.
(function readGoogleReturn() {
  if (!location.hash.includes("state=focus-google")) return;
  const q = new URLSearchParams(location.hash.slice(1));
  if (q.get("access_token")) localStorage.setItem("focus-google-token", JSON.stringify({ token: q.get("access_token"), exp: Date.now() + (Number(q.get("expires_in")) || 3600) * 1000 }));
  history.replaceState(null, "", location.pathname + location.search);
  sessionStorage.setItem("focus-open-calendar", "1"); // go straight back to the calendar
})();
async function gFetch(path, opts = {}) {
  const t = gToken();
  if (!t) throw Object.assign(new Error("Connect Google Calendar to see it here."), { needsConnect: true });
  const r = await fetch("https://www.googleapis.com/calendar/v3/calendars/" + encodeURIComponent(gCalendarId()) + path,
    { ...opts, headers: { Authorization: "Bearer " + t, "Content-Type": "application/json", ...(opts.headers || {}) } });
  if (r.status === 401) { localStorage.removeItem("focus-google-token"); throw Object.assign(new Error("Google's sign-in has run out. Connect again."), { needsConnect: true }); }
  if (!r.ok && r.status !== 204) throw new Error("Google Calendar: " + r.status + " " + ((await r.text()).slice(0, 120)));
  return r.status === 204 ? null : r.json();
}
/** Events from your Google calendar between two dates (repeating events come as separate days). */
async function gEvents(from, to) {
  const out = []; let page = "";
  do {
    const j = await gFetch("/events?" + new URLSearchParams({ timeMin: from.toISOString(), timeMax: to.toISOString(), singleEvents: "true", orderBy: "startTime", maxResults: "250", ...(page ? { pageToken: page } : {}) }));
    for (const e of j.items || []) {
      if (e.status === "cancelled") continue;
      const allDay = !!(e.start && e.start.date);
      out.push({ id: e.id, title: e.summary || "(No title)", allDay, kind: "meeting",
                 start: allDay ? parseYmd(e.start.date) : new Date(e.start.dateTime), end: allDay ? parseYmd(e.end.date) : new Date(e.end.dateTime), raw: e });
    }
    page = j.nextPageToken || "";
  } while (page);
  return out;
}
const zone = () => Intl.DateTimeFormat().resolvedOptions().timeZone;
function gBody(f) {
  if (f.allDay) return { summary: f.title, start: { date: f.date }, end: { date: ymd(addDays(parseYmd(f.date), 1)) } };
  return { summary: f.title, start: { dateTime: `${f.date}T${f.from}:00`, timeZone: zone() }, end: { dateTime: `${f.date}T${f.to}:00`, timeZone: zone() } };
}

// ---------- The calendar: Day and Month ----------

let calDay = null, calMonth = null;
const calKind = () => localStorage.getItem("focus-cal-kind") === "month" ? "month" : "day";
const timeText = d => d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });

/** Your work calendar (shared by Focus on your computer), and optionally your plan, for one day. */
async function workDay(day) {
  const shared = await readFile("calendar-share.json");
  const items = (Array.isArray(shared) ? shared : [])
    .map(e => ({ title: e.Title || "Busy", start: parseDate(e.Start), end: parseDate(e.End), allDay: !!e.AllDay, busy: e.Busy !== false, kind: "meeting" }))
    .filter(e => e.start && e.end && e.start < addDays(day, 1) && e.end > day);
  if (calShowPlan()) {
    const plan = await readFile(`plans/${ymd(day)}.json`);
    if (plan && !Array.isArray(plan) && Array.isArray(plan.Blocks)) {
      const [h, m] = String(plan.DayStart || "06:45").split(":").map(Number);
      let t = new Date(day); t.setHours(h || 0, m || 0, 0, 0);
      for (const b of plan.Blocks) {
        const start = new Date(t), end = new Date(t.getTime() + (b.Minutes || 0) * 60000); t = end;
        if (b.Kind === "Meeting") continue; // meetings come from the calendar itself
        items.push({ title: b.Name || (b.Kind === "Focus" ? "Focus" : "Break"), start, end, kind: b.Kind === "Focus" ? "focus" : "break",
                     tasks: (b.Tasks || []).filter(x => x && x.Text).map(x => x.Text) });
      }
    }
  }
  return { items, note: Array.isArray(shared) && shared.length ? "Meetings come from Focus on your computer (updated every 15 minutes while it's on)."
                                                                : "No meetings shared yet. Focus on your computer shares your calendar once it's connected." };
}

function dayHtml(items, day, editable) {
  items.sort((a, b) => (a.allDay ? -1 : 0) - (b.allDay ? -1 : 0) || a.start - b.start || (a.kind === "meeting" ? -1 : 1));
  const now = new Date();
  let html = "", nowShown = !sameDay(day, now);
  items.forEach((it, i) => {
    if (!nowShown && !it.allDay && it.start > now) { html += `<div class="nowline"><span>Now ${timeText(now)}</span></div>`; nowShown = true; }
    html += `<div class="cal ${it.kind}${it.busy === false ? " free" : ""}${editable ? " evrow" : ""}" data-i="${i}"><div class="when">${it.allDay ? "All day" : `${timeText(it.start)}<br><small>${timeText(it.end)}</small>`}</div>` +
      `<div class="what">${esc(it.title)}${it.tasks && it.tasks.length ? `<small>${it.tasks.map((x, n) => `${n + 1}. ${esc(x)}`).join("<br>")}</small>` : ""}</div></div>`;
  });
  if (!nowShown) html += `<div class="nowline"><span>Now ${timeText(now)}</span></div>`;
  return html;
}

/** Opens the calendar: Day or Month (your choice is remembered). Personal shows your Google calendar; Tasks your work one. */
async function openCalendar(day, kind) {
  if (kind) localStorage.setItem("focus-cal-kind", kind);
  calDay = startOfDay(day || calDay || new Date());
  const personal = mode === "personal", month = calKind() === "month";
  if (month) calMonth = new Date(calDay.getFullYear(), calDay.getMonth(), 1);
  const card = $("sheetCard");
  const title = month ? calMonth.toLocaleDateString([], { month: "long", year: "numeric" })
                      : sameDay(calDay, today()) ? "Today" : calDay.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" });
  card.innerHTML = `<div class="seg" style="margin-bottom:10px"><button type="button" id="cDay" aria-pressed="${!month}">Day</button><button type="button" id="cMonth" aria-pressed="${month}">Month</button></div>
    <div class="calhead"><button type="button" class="icon" id="cPrev" aria-label="Previous">‹</button><b>${esc(title)}</b><button type="button" class="icon" id="cNext" aria-label="Next">›</button></div>
    <div class="calbar"><button type="button" id="cToday">Today</button>
      ${personal ? (month ? "" : `<button type="button" id="cAdd">+ Event</button>`) : (month ? "" : `<label class="switch"><input type="checkbox" id="cPlan" ${calShowPlan() ? "checked" : ""}> Show my plan</label>`)}</div>
    <div id="cList" class="callist"><p class="empty">Loading...</p></div>
    <div class="hint" id="cNote"></div>
    <div class="sheetbar"><span></span><button type="button" id="cClose">Close</button></div>`;
  $("cDay").onclick = () => openCalendar(calDay, "day");
  $("cMonth").onclick = () => openCalendar(calDay, "month");
  $("cPrev").onclick = () => openCalendar(month ? new Date(calMonth.getFullYear(), calMonth.getMonth() - 1, 1) : addDays(calDay, -1));
  $("cNext").onclick = () => openCalendar(month ? new Date(calMonth.getFullYear(), calMonth.getMonth() + 1, 1) : addDays(calDay, 1));
  $("cToday").onclick = () => openCalendar(new Date());
  if ($("cPlan")) $("cPlan").onchange = e => { localStorage.setItem("focus-cal-plan", e.target.checked ? "on" : "off"); openCalendar(calDay); };
  if ($("cAdd")) $("cAdd").onclick = () => editEvent(null, calDay);
  $("cClose").onclick = closeSheet;
  openSheet();
  const shown = calDay, shownKind = calKind();
  const still = () => shown === calDay && shownKind === calKind() && !$("sheet").classList.contains("hidden");
  try {
    if (month) {
      // Six weeks from the Monday before the 1st, each day marked with dots for its events.
      const first = startOfWeek(calMonth), last = addDays(first, 42);
      let events = [];
      if (personal) events = await gEvents(first, last);
      else { const shared = await readFile("calendar-share.json"); events = (Array.isArray(shared) ? shared : []).map(e => ({ start: parseDate(e.Start), end: parseDate(e.End) })).filter(e => e.start); }
      if (!still()) return;
      let html = `<div class="month">${["M", "T", "W", "T", "F", "S", "S"].map(d => `<div class="dow">${d}</div>`).join("")}`;
      for (let n = 0; n < 42; n++) {
        const d = addDays(first, n);
        const count = events.filter(e => e.start < addDays(d, 1) && (e.end || e.start) > d).length;
        html += `<button type="button" data-day="${ymd(d)}" class="${d.getMonth() !== calMonth.getMonth() ? "other" : ""}${sameDay(d, today()) ? " today" : ""}">${d.getDate()}<span class="marks">${"•".repeat(Math.min(count, 3))}</span></button>`;
      }
      $("cList").innerHTML = html + "</div>";
      $("cList").querySelectorAll("[data-day]").forEach(b => b.onclick = () => openCalendar(parseYmd(b.dataset.day), "day"));
      $("cNote").textContent = personal ? "Your Google calendar. Tap a day to see it." : "Your work calendar from Focus. Tap a day to see it.";
      return;
    }
    if (personal) {
      const items = await gEvents(shown, addDays(shown, 1));
      if (!still()) return;
      $("cList").innerHTML = items.length ? dayHtml(items, shown, true) : `<p class="empty">Nothing on your Google calendar.</p>`;
      $("cList").querySelectorAll(".evrow").forEach(row => row.onclick = () => editEvent(items[+row.dataset.i], shown));
      $("cNote").textContent = "Your Google calendar: tap an event to change it. Changes show in Google too.";
    } else {
      const { items, note } = await workDay(shown);
      if (!still()) return;
      $("cList").innerHTML = items.length ? dayHtml(items, shown, false) : `<p class="empty">Nothing on the calendar${calShowPlan() ? " or planned" : ""}.</p>`;
      $("cNote").textContent = note;
    }
  } catch (e) {
    if (!still()) return;
    $("cList").innerHTML = e.needsConnect
      ? `<p class="empty">${esc(e.message)}</p><button type="button" class="primary" id="cConnect" style="width:100%">${gClientId() ? "Connect Google Calendar" : "Set up Google Calendar"}</button>`
      : `<p class="empty">${esc(e.message)}</p>`;
    if ($("cConnect")) $("cConnect").onclick = () => gClientId() ? gConnect() : openMenu();
  }
}

/** Adding or changing an event on your Google calendar (it changes in Google too). */
function editEvent(ev, day) {
  const d = ev ? ev.start : day;
  const pad = n => String(n).padStart(2, "0");
  const hm = x => `${pad(x.getHours())}:${pad(x.getMinutes())}`;
  const from = ev && !ev.allDay ? hm(ev.start) : "09:00", to = ev && !ev.allDay ? hm(ev.end) : "10:00";
  const card = $("sheetCard");
  card.innerHTML = `<div class="evform"><label>${ev ? "Event" : "New event"}</label>
    <input id="eTitle" value="${esc(ev ? ev.title : "")}" placeholder="What" autocomplete="off">
    <label>Day</label><input id="eDate" type="date" value="${ymd(d)}">
    <label class="switch" style="margin-top:12px"><input type="checkbox" id="eAll" ${ev && ev.allDay ? "checked" : ""}> All day</label>
    <div class="two" id="eTimes"><div><label>From</label><input id="eFrom" type="time" value="${from}"></div><div><label>To</label><input id="eTo" type="time" value="${to}"></div></div></div>
    <div class="sheetbar">${ev ? `<button type="button" id="eDelete">Delete</button>` : "<span></span>"}<span><button type="button" id="eCancel">Cancel</button> <button type="button" class="primary" id="eSave">Save</button></span></div>`;
  const times = () => $("eTimes").classList.toggle("hidden", $("eAll").checked);
  times(); $("eAll").onchange = times;
  $("eCancel").onclick = () => openCalendar(day, "day");
  $("eSave").onclick = async () => {
    const f = { title: $("eTitle").value.trim() || "(No title)", date: $("eDate").value, allDay: $("eAll").checked, from: $("eFrom").value, to: $("eTo").value };
    if (!f.date || (!f.allDay && f.to <= f.from)) { toast("Check the day and times."); return; }
    try {
      if (ev) await gFetch("/events/" + encodeURIComponent(ev.id), { method: "PATCH", body: JSON.stringify(gBody(f)) });
      else await gFetch("/events", { method: "POST", body: JSON.stringify(gBody(f)) });
      toast(ev ? "Event changed" : "Event added");
      openCalendar(parseYmd(f.date), "day");
    } catch (e) { toast(e.message); }
  };
  if ($("eDelete")) $("eDelete").onclick = async () => {
    if (!confirm("Delete this event from your Google calendar?")) return;
    try { await gFetch("/events/" + encodeURIComponent(ev.id), { method: "DELETE" }); toast("Event deleted"); openCalendar(day, "day"); }
    catch (e) { toast(e.message); }
  };
  $("eTitle").focus();
}

// ---------- Menu and themes ----------

const THEMES = [["classic", "Classic"], ["mono", "Monochrome"], ["calm", "Calm"]];
/** Personal's accent colors: [name, light mode, dark mode]. */
const ACCENTS = [["green", "Green", "#2F9E5E", "#5CC88A"], ["teal", "Teal", "#13897C", "#4CC3B5"], ["blue", "Blue", "#246FE0", "#6FA3F0"],
                 ["violet", "Violet", "#7C5CD6", "#A88BF5"], ["amber", "Amber", "#C77D12", "#E8A84A"], ["red", "Red", "#DC4C3E", "#F07A6E"]];
const darkMode = window.matchMedia("(prefers-color-scheme: dark)");
const valid = (k, list) => list.some(t => t[0] === k);

/** Your theme on the Tasks page; Personal's own theme and accent on the Personal page. The phone's bars follow. */
function applyTheme() {
  const root = document.documentElement;
  if (mode === "personal") {
    const k = localStorage.getItem("focus-ptheme") || "calm";
    root.dataset.theme = valid(k, THEMES) ? k : "calm";
    const a = ACCENTS.find(x => x[0] === (localStorage.getItem("focus-paccent") || "green")) || ACCENTS[0];
    const hex = darkMode.matches ? a[3] : a[2];
    root.style.setProperty("--accent", hex);
    root.style.setProperty("--tint", hex + (darkMode.matches ? "33" : "22"));
  } else {
    const k = localStorage.getItem("focus-theme") || "classic";
    root.dataset.theme = valid(k, THEMES) ? k : "classic";
    root.style.removeProperty("--accent");
    root.style.removeProperty("--tint");
  }
  // The phone's own top and bottom bars take the page's background color (they used to stay red).
  const bg = getComputedStyle(root).getPropertyValue("--bg").trim() || "#FAF9F7";
  document.querySelectorAll('meta[name="theme-color"]').forEach(m => m.setAttribute("content", bg));
}
darkMode.addEventListener("change", applyTheme);

function openMenu() {
  const mainTheme = localStorage.getItem("focus-theme") || "classic";
  const pTheme = localStorage.getItem("focus-ptheme") || "calm";
  const pAccent = localStorage.getItem("focus-paccent") || "green";
  const seg = (items, current, kind) => `<div class="seg">${items.map(([k, n]) => `<button type="button" data-kind="${kind}" data-k="${k}" aria-pressed="${k === current}">${n}</button>`).join("")}</div>`;
  const card = $("sheetCard");
  card.innerHTML = `<label>Theme</label>${seg(THEMES, mainTheme, "theme")}
    <label style="margin-top:14px">Personal page</label>${seg(THEMES, pTheme, "ptheme")}
    <div class="seg" style="margin-top:6px">${ACCENTS.map(([k, n, light]) => `<button type="button" data-kind="paccent" data-k="${k}" aria-pressed="${k === pAccent}"><span style="display:inline-block;width:10px;height:10px;border-radius:5px;background:${light};margin-right:5px"></span>${n}</button>`).join("")}</div>
    <div class="hint">Light or dark follows your phone. Swipe left for Tasks, right for Personal.</div>
    <label style="margin-top:14px">Personal opens on</label>${seg([["all", "All"], ["today", "Today"]], openingView("personal"), "openpersonal")}
    <label style="margin-top:10px">Tasks opens on</label>${seg([["all", "All"], ["today", "Today"]], openingView("main"), "openmain")}
    <label style="margin-top:14px">Google Calendar (Personal's calendar)</label>
    <input id="mGClient" placeholder="Google client ID (see SETUP.md)" value="${esc(gClientId())}" autocomplete="off">
    <input id="mGCal" placeholder="Calendar ID (leave as primary for your main calendar)" value="${esc(gCalendarId())}" autocomplete="off" style="margin-top:6px">
    <div class="hint">${gToken() ? "Connected." : "Not connected."} Events you add, change or delete in Personal's calendar change in Google too.</div>
    <div style="display:flex;gap:8px;margin-top:6px"><button type="button" id="mGConnect" style="flex:1">${gToken() ? "Reconnect" : "Connect"}</button>${gToken() ? `<button type="button" id="mGOff" style="flex:1">Disconnect</button>` : ""}</div>
    <button type="button" id="mRefresh" style="width:100%;margin-top:12px">Refresh now</button>
    <button type="button" id="mExport" style="width:100%;margin-top:12px">Export or print a list</button>
    <div class="sheetbar"><button type="button" id="mSignout">Sign out</button><button type="button" id="mClose">Close</button></div>`;
  const keys = { theme: "focus-theme", ptheme: "focus-ptheme", paccent: "focus-paccent", openpersonal: "focus-open-personal", openmain: "focus-open-main" };
  const saveGoogle = () => { localStorage.setItem("focus-google-client", $("mGClient").value.trim()); localStorage.setItem("focus-google-calendar", $("mGCal").value.trim() || "primary"); };
  $("mGClient").onchange = saveGoogle; $("mGCal").onchange = saveGoogle;
  $("mGConnect").onclick = () => { saveGoogle(); gConnect(); };
  if ($("mGOff")) $("mGOff").onclick = () => { gDisconnect(); openMenu(); };
  $("mRefresh").onclick = () => { closeSheet(); reload(); };
  card.querySelectorAll(".seg button").forEach(b => b.onclick = () => { localStorage.setItem(keys[b.dataset.kind], b.dataset.k); applyTheme(); openMenu(); });
  $("mSignout").onclick = () => { closeSheet(); client.logoutRedirect({ account }); };
  $("mExport").onclick = openExport;
  $("mClose").onclick = closeSheet;
  openSheet();
}
applyTheme();

// ---------- Adding a task: the + button ----------

function openAdd() {
  const card = $("sheetCard");
  card.innerHTML = `<form class="addsheet" id="aForm"><label>${mode === "personal" ? "Add a personal task" : "Add a task"}</label>
    <input id="aText" autocomplete="off" enterkeyhint="done" placeholder="Call Sam friday at 3pm #Client @calls !">
    <small>Add as many as you like; each one is saved when you press Enter.</small>
    <div class="sheetbar"><span></span><span><button type="button" id="aClose">Done</button> <button type="submit" class="primary">Add</button></span></div></form>`;
  $("aForm").onsubmit = e => { e.preventDefault(); const v = $("aText").value.trim(); if (!v) return; $("aText").value = ""; add(v); $("aText").focus(); };
  $("aClose").onclick = closeSheet;
  openSheet();
  setTimeout(() => $("aText").focus(), 60);
}

// ---------- Pull down to refresh ----------

(function pullToRefresh() {
  let y0 = null, x0 = 0, pulled = false;
  document.addEventListener("touchstart", e => {
    y0 = window.scrollY <= 0 && $("sheet").classList.contains("hidden") && e.touches.length === 1 ? e.touches[0].clientY : null;
    x0 = e.touches[0].clientX; pulled = false;
  }, { passive: true });
  document.addEventListener("touchmove", e => {
    if (y0 === null) return;
    const dy = e.touches[0].clientY - y0, dx = Math.abs(e.touches[0].clientX - x0);
    pulled = dy > 90 && dx < 50;
    $("ptr").classList.toggle("hidden", !pulled);
  }, { passive: true });
  document.addEventListener("touchend", () => {
    $("ptr").classList.add("hidden");
    if (pulled && account) reload();
    y0 = null; pulled = false;
  }, { passive: true });
})();

// ---------- Tasks and Personal: swipe sideways (or tap the dots) ----------

const pages = { main: { bin: [], resolved: [], projects: [], view: openingView("main") }, personal: { bin: [], resolved: [], projects: [], view: openingView("personal") } };
function switchPage(to) {
  if (to === mode || busy) return;
  pages[mode] = { bin, resolved, projects, view: view.startsWith("project:") ? "all" : view };
  mode = to;
  ({ bin, resolved, projects, view } = pages[mode]);
  chip = "all"; search = "";
  if ($("search")) { $("search").value = ""; $("search").classList.add("hidden"); }
  document.querySelector("header h1").textContent = mode === "personal" ? "Personal" : "Tasks";
  document.querySelectorAll("#pageDots span").forEach((s, i) => s.classList.toggle("on", (i === 0) === (mode === "personal")));
  applyTheme();
  render();
  reload();
}
(function swipes() {
  let x0 = null, y0 = null, t0 = 0;
  document.addEventListener("touchstart", e => {
    if (!$("sheet").classList.contains("hidden") || e.touches.length !== 1) { x0 = null; return; }
    const tag = (e.target.tagName || "").toLowerCase();
    if (tag === "input" || tag === "textarea") { x0 = null; return; }
    x0 = e.touches[0].clientX; y0 = e.touches[0].clientY; t0 = Date.now();
  }, { passive: true });
  document.addEventListener("touchend", e => {
    if (x0 === null) return;
    const dx = e.changedTouches[0].clientX - x0, dy = e.changedTouches[0].clientY - y0;
    x0 = null;
    if (Math.abs(dx) < 70 || Math.abs(dx) < Math.abs(dy) * 2 || Date.now() - t0 > 700) return;
    // Personal is on the left, Tasks on the right: swipe left to go to Tasks, right to come back.
    if (dx < 0 && mode === "personal") switchPage("main");
    else if (dx > 0 && mode === "main") switchPage("personal");
  }, { passive: true });
  document.querySelectorAll("#pageDots span").forEach((s, i) => s.onclick = () => switchPage(i === 0 ? "personal" : "main"));
})();
// Opening on Personal: its title, entry bar and look from the start.
document.querySelector("header h1").textContent = "Personal";
applyTheme();

// ---------- Export and print (the same lists as on the computer) ----------

/** The lists you can print, like the computer's Export menu. */
function exportChoices() {
  const d = startOfDay(new Date()), tomorrow = addDays(d, 1), weekEnd = addDays(d, (7 - d.getDay()) % 7 || 7);
  const open = bin.filter(t => !isSnoozed(t, d));
  const isToday = t => isDueToday(t, d) || isOverdue(t, d) || isPlannedToday(t, d);
  const one = (header, items) => [{ header, items: sortTasks(items, d) }];
  const viewNames = { today: "Today", all: "All tasks", upcoming: "Upcoming", undated: "Undated", done: "Done" };
  const list = [
    ["Today", () => [
      { header: "Overdue", items: sortTasks(open.filter(t => isOverdue(t, d)), d) },
      { header: "Due today", items: sortTasks(open.filter(t => isDueToday(t, d)), d) },
      { header: "Picked for today", items: sortTasks(open.filter(t => isPlannedToday(t, d) && !isDueToday(t, d) && !isOverdue(t, d)), d) }]],
    ["Tomorrow", () => one("Tomorrow", bin.filter(t => sameDay(dueOf(t), tomorrow) || t.PlannedFor === ymd(tomorrow)))],
    ["This week", () => one("Due by " + weekEnd.toLocaleDateString(undefined, { weekday: "long" }), open.filter(t => { const due = dueOf(t); return due && due <= weekEnd; }))],
    ["Upcoming", () => {
      const byDay = {};
      for (const t of bin.filter(t => dueOf(t) && dueOf(t) > d)) (byDay[ymd(dueOf(t))] ??= []).push(t);
      return Object.keys(byDay).sort().map(k => ({ header: parseYmd(k).toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" }), items: sortTasks(byDay[k], d) }));
    }],
    ["Overdue", () => one("Overdue", open.filter(t => isOverdue(t, d)))],
    ["Undated", () => one("Undated", open.filter(isUndated))],
    ["Waiting on", () => one("Waiting on", open.filter(t => t.Status === "waiting"))],
    ["All tasks", () => {
      const groups = projects.map(p => ({ header: p.Name, items: sortTasks(open.filter(t => t.Project === p.Id), d) }));
      groups.push({ header: "No project", items: sortTasks(open.filter(t => !t.Project || !projects.some(p => p.Id === t.Project)), d) });
      return groups.filter(g => g.items.length);
    }],
    ["Done today", () => [{ header: "Done today", done: true, items: resolved.filter(r => (parseDate(r.ResolvedAt) || 0) >= d) }]],
    ["Done this week", () => { const monday = startOfWeek(d); return [{ header: "Done this week", done: true, items: resolved.filter(r => (parseDate(r.ResolvedAt) || 0) >= monday) }]; }],
  ];
  if (view !== "done" && !view.startsWith("project:")) list.unshift([`This view (${viewNames[view] || "Tasks"})`, () => list.find(x => x[0] === (viewNames[view] === "All tasks" ? "All tasks" : viewNames[view]))?.[1]() || one("Tasks", open)]);
  const projectLists = projects.map(p => [p.Name, () => one(p.Name, open.filter(t => t.Project === p.Id))]);
  const tags = [...new Set(open.flatMap(t => t.Tags || []))].sort();
  const tagLists = tags.map(g => ["@" + g, () => one("@" + g, open.filter(t => (t.Tags || []).includes(g)))]);
  return { list, projectLists, tagLists };
}

function openExport() {
  const { list, projectLists, tagLists } = exportChoices();
  const look = localStorage.getItem("focus-printlook") || "default";
  const seg = `<div class="seg">${[["default", "Default"], ["compact", "Compact"], ["paper", "Paper"]].map(([k, n]) =>
    `<button type="button" data-look="${k}" aria-pressed="${k === look}">${n}</button>`).join("")}</div>`;
  const rows = (items, offset) => items.map(([name], i) => `<button type="button" class="exportRow" data-i="${i + offset}">${esc(name)}</button>`).join("");
  const all = [...list, ...projectLists, ...tagLists];
  const card = $("sheetCard");
  card.innerHTML = `<label>Print or save as a PDF</label>${seg}
    <div class="exportList">${rows(list, 0)}
    ${projectLists.length ? `<label style="margin-top:12px">Projects</label>${rows(projectLists, list.length)}` : ""}
    ${tagLists.length ? `<label style="margin-top:12px">Tags</label>${rows(tagLists, list.length + projectLists.length)}` : ""}</div>
    <div class="sheetbar"><button type="button" id="xClose">Close</button></div>`;
  card.querySelectorAll("[data-look]").forEach(b => b.onclick = () => { localStorage.setItem("focus-printlook", b.dataset.look); openExport(); });
  card.querySelectorAll(".exportRow").forEach(b => b.onclick = () => { const [name, build] = all[+b.dataset.i]; closeSheet(); printGroups(name, build()); });
  $("xClose").onclick = closeSheet;
  openSheet();
}

/** Lays the list out like the computer's printout (logo, title and date, circles to tick, details, due dates), then opens print. */
function printGroups(title, groups) {
  const d = startOfDay(new Date());
  const look = localStorage.getItem("focus-printlook") || "default";
  const total = groups.reduce((n, g) => n + g.items.length, 0);
  const date = new Date();
  const fmtDue = due => due.getFullYear() === d.getFullYear() ? due.toLocaleDateString(undefined, { month: "short", day: "numeric" }) : due.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
  const row = (t, done) => {
    const proj = projects.find(p => p.Id === t.Project);
    const detail = [proj ? proj.Name : "", ...(t.Tags || []).map(g => "@" + g)].filter(Boolean).join("   ");
    const due = !done && dueOf(t);
    const high = !done && t.Priority === "High";
    return `<div class="p-row"><span class="p-circle${done ? " p-done" : ""}">${done ? "✓" : ""}</span>
      <div class="p-text"><div class="${done ? "p-struck" : ""}">${esc(t.Text)}</div>${detail ? `<div class="p-detail">${esc(detail)}</div>` : ""}
      ${(t.Subtasks || []).map(s => `<div class="p-step"><span class="p-circle p-small${s.Done ? " p-done" : ""}">${s.Done ? "✓" : ""}</span>${esc(s.Text)}</div>`).join("")}</div>
      <div class="p-side">${high ? '<span class="p-high">High</span> ' : ""}${due ? `<span class="${due < d ? "p-high" : ""}">${fmtDue(due)}</span>` : ""}</div></div>`;
  };
  const shown = groups.filter(g => g.items.length || groups.length === 1);
  const html = `<div class="p-page p-${look}">
    <div class="p-head"><img src="icon-192.png" alt=""><div class="p-names"><div class="p-brand">FOCUS</div><div class="p-title">${esc(title)}</div></div>
      <div class="p-when"><div>${date.toLocaleDateString(undefined, { weekday: "long" })}</div><b>${date.toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" })}</b></div></div>
    <div class="p-rule"></div>
    <div class="p-summary">${total === 0 ? "Nothing here right now." : total === 1 ? "1 task" : total + " tasks"}</div>
    ${shown.map(g => `${shown.length > 1 || g.header !== title ? `<div class="p-group">${esc(g.header)} <span>${g.items.length}</span></div>` : ""}
      ${g.items.length ? g.items.map(t => row(t, !!g.done)).join("") : '<div class="p-detail">Nothing here right now.</div>'}`).join("")}
    <div class="p-notes">Notes</div><div class="p-line"></div><div class="p-line"></div><div class="p-line"></div><div class="p-line"></div>
    <div class="p-foot">Printed from Focus, ${date.toLocaleString(undefined, { month: "long", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" })}</div></div>`;
  let area = $("printArea");
  if (!area) { area = document.createElement("div"); area.id = "printArea"; document.body.appendChild(area); }
  area.innerHTML = html;
  setTimeout(() => window.print(), 150); // let the logo load first
}

// ---------- Wiring up ----------

$("signin").onclick = () => client.loginRedirect({ scopes: SCOPES });
$("menuBtn").onclick = openMenu;
$("projectsBtn").onclick = openProjects;
$("calBtn").onclick = () => openCalendar(new Date());
// Back from Google's sign-in: open the calendar again once the app is ready.
if (sessionStorage.getItem("focus-open-calendar")) { sessionStorage.removeItem("focus-open-calendar"); setTimeout(() => openCalendar(new Date()), 900); }
$("projectBack").onclick = () => { view = "all"; render(); };
$("sheet").addEventListener("click", e => { if (e.target.id === "sheet") closeSheet(); });
$("addBtn").onclick = openAdd;
$("filterBtn").onclick = () => { filtersOpen = !filtersOpen; if (!filtersOpen) chip = "all"; render(); };
$("searchBtn").onclick = () => {
  const box = $("search"); box.classList.toggle("hidden");
  if (!box.classList.contains("hidden")) box.focus(); else { search = ""; box.value = ""; render(); }
};
$("search").oninput = e => { search = e.target.value.trim(); render(); };
$("tabs").querySelectorAll("button").forEach(b => b.onclick = () => { view = b.dataset.view; render(); });
// Get the latest whenever you come back to the page (unless you're in the middle of editing).
document.addEventListener("visibilitychange", () => { if (!document.hidden && account && $("sheet").classList.contains("hidden")) reload(); });

// Lets your phone install the page as an app (and open it offline to the last copy).
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => { /* still works as a normal page */ });

$("boot").classList.add("hidden");
start();
