/**
 * Venue Logger - clickable dialog.
 *
 * Builds one Log row from four slots (Callsign, Department, Incident,
 * Location) plus optional notes. Buttons come from the Inputs sheet, sent
 * over by commands.js. Typing shorthand with no buttons picked still works
 * exactly like before.
 *
 * Messages to commands.js (JSON strings):
 *   { kind: "ready" }                      - send me the buttons
 *   { kind: "text",  text: "l1s ps ed14" } - old shorthand line
 *   { kind: "entry", fields: [callsign, department, location, details] }
 * Messages from commands.js:
 *   { kind: "menuChunk", id, i, n, data }  - button data, split into chunks
 *   { kind: "status", ok, text }           - result of a log write
 */

const STAGES = [
  { key: "callsign",   name: "Callsign",   prompt: "Pick a callsign" },
  { key: "department", name: "Department", prompt: "Pick a department" },
  { key: "incident",   name: "Incident",   prompt: "Pick an incident" },
  { key: "location",   name: "Location",   prompt: "Pick a location" },
];
const RECENTS_KEY = "venueLogger.recents";
const RECENTS_MAX = 6;
const SEND_TIMEOUT_MS = 12000;
// Top-level location groups shown first, in this order. Any other groups
// follow in the order they first appear on the Inputs sheet.
const LOCATION_GROUP_ORDER = [
  "External", "Internal", "Blocks",
  "Basement", "Ground Floor", "Level 1", "Level 2", "Level 3", "Level 4",
  "Catwalk", "Block", "Toilets",
];

let menu = { callsign: [], department: [], incident: [], location: [] };
let deptColour = {};   // lower-case department label -> colour
let locTree = null;

let sel = emptySelection();
let skipped = {};
let stage = "callsign";
let locPath = [];      // names of the location groups drilled into
let busy = false;
let busyTimer = null;
let pendingRecent = null;
let menuLoaded = false;
const chunkBuffer = {};

let transport = null;  // { send(obj) }

const $ = (id) => document.getElementById(id);

/* ==========================================================================
 * Start-up
 * ========================================================================*/

function start(t) {
  transport = t;
  $("logInput").addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); logEntry(); }
    if (e.key === "Escape") { e.preventDefault(); clearAll(); }
  });
  $("logBtn").addEventListener("click", logEntry);
  $("clearBtn").addEventListener("click", clearAll);
  $("refreshBtn").addEventListener("click", requestMenu);
  $("skipBtn").addEventListener("click", skipStage);
  $("backBtn").addEventListener("click", () => { locPath.pop(); render(); });
  render();
  requestMenu();
  $("logInput").focus();
}

if (window.VL_DEMO) {
  // Browser preview: no Excel, fake transport supplied by the preview page.
  start(window.VL_DEMO.transport(onParentMessage));
} else {
  Office.onReady(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      onParentMessage(arg.message);
    });
    start({ send: (obj) => Office.context.ui.messageParent(JSON.stringify(obj)) });
  });
}

function requestMenu() {
  setStatus("Loading buttons…", "busy");
  transport.send({ kind: "ready" });
  // The dialog can load before commands.js has attached its listener, so
  // ask once more if nothing has arrived.
  setTimeout(() => { if (!menuLoaded) transport.send({ kind: "ready" }); }, 3000);
}

/* ==========================================================================
 * Messages from commands.js
 * ========================================================================*/

function onParentMessage(raw) {
  let msg;
  try { msg = JSON.parse(raw); } catch (e) { msg = { kind: "status", ok: !/^error/i.test(raw), text: raw }; }

  if (msg.kind === "menuChunk") {
    const buf = chunkBuffer[msg.id] || (chunkBuffer[msg.id] = []);
    buf[msg.i] = msg.data;
    if (buf.filter((x) => x !== undefined).length === msg.n) {
      delete chunkBuffer[msg.id];
      try {
        ingestMenu(JSON.parse(buf.join("")));
      } catch (e) {
        setStatus("Error: couldn't read the button data. Click Reload buttons.", "err");
      }
    }
    return;
  }

  if (msg.kind === "status") {
    clearTimeout(busyTimer);
    busy = false;
    setStatus(msg.text, msg.ok ? "ok" : "err");
    $("logInput").disabled = false;
    $("logBtn").disabled = false;
    if (msg.ok) {
      if (pendingRecent) addRecent(pendingRecent);
      pendingRecent = null;
      resetEntry();
    }
    $("logInput").focus();
  }
}

/* ==========================================================================
 * Button data
 * ========================================================================*/

/** rows: [code, label, type, group, colour][] */
function ingestMenu(rows) {
  menu = { callsign: [], department: [], incident: [], location: [] };
  for (const r of rows) {
    const type = (r[2] || "").toLowerCase();
    if (!menu[type]) continue;
    menu[type].push({ code: r[0] || "", label: r[1] || "", group: r[3] || "", colour: r[4] || "", type });
  }

  deptColour = {};
  for (const d of menu.department) deptColour[d.label.toLowerCase()] = d.colour || "";
  for (const inc of menu.incident) {
    if (!inc.colour) inc.colour = deptColour[inc.group.toLowerCase()] || "";
  }
  locTree = buildTree(menu.location);
  menuLoaded = true;

  // Re-point any current picks at the fresh items so labels/colours update.
  for (const s of STAGES) {
    if (sel[s.key]) sel[s.key] = menu[s.key].find((i) => i.label === sel[s.key].label) || sel[s.key];
  }

  const total = rows.length;
  setStatus(total ? "Ready." : "No buttons found. Fill in the Button type column on Inputs.", total ? "ok" : "err");
  render();
  renderRecents();
}

/**
 * Location groups are paths like "Blocks/200s". Several paths can be given
 * with ";" (e.g. "Level 2;Toilets") so one location appears in both places.
 */
function buildTree(items) {
  const root = newNode("");
  for (const item of items) {
    const paths = item.group.split(";").map((s) => s.trim()).filter(Boolean);
    if (paths.length === 0) paths.push("");
    for (const p of paths) {
      let node = root;
      for (const part of p.split("/").map((s) => s.trim()).filter(Boolean)) {
        if (!node.childMap[part]) {
          const child = newNode(part);
          child.colour = item.colour;
          node.childMap[part] = child;
          node.children.push(child);
        }
        node = node.childMap[part];
      }
      node.items.push(item);
    }
  }
  const rank = (name) => {
    const i = LOCATION_GROUP_ORDER.findIndex((g) => g.toLowerCase() === name.toLowerCase());
    return i === -1 ? LOCATION_GROUP_ORDER.length : i;
  };
  root.children = root.children
    .map((c, i) => ({ c, i }))
    .sort((a, b) => rank(a.c.name) - rank(b.c.name) || a.i - b.i)
    .map((x) => x.c);
  return root;
}

function newNode(name) {
  return { name, colour: "", children: [], childMap: {}, items: [] };
}

function countItems(node) {
  return node.items.length + node.children.reduce((n, c) => n + countItems(c), 0);
}

function currentLocNode() {
  let node = locTree;
  for (const name of locPath) {
    if (!node || !node.childMap[name]) return locTree;
    node = node.childMap[name];
  }
  return node;
}

function incidentsFor(dept) {
  if (!dept) return [];
  const d = dept.label.toLowerCase();
  return menu.incident.filter((i) => i.group.toLowerCase() === d);
}

/* ==========================================================================
 * Picking, skipping, moving between stages
 * ========================================================================*/

function emptySelection() {
  return { callsign: null, department: null, incident: null, location: null };
}

function applicable(key) {
  if (key === "incident") return incidentsFor(sel.department).length > 0;
  return true;
}

function pick(key, item) {
  sel[key] = item;
  skipped[key] = false;
  if (key === "department" && sel.incident && incidentsFor(item).indexOf(sel.incident) === -1) {
    sel.incident = null;
  }
  if (key === "location") locPath = [];
  advanceFrom(key);
}

function skipStage() {
  if (stage === "done") return;
  sel[stage] = null;
  skipped[stage] = true;
  if (stage === "location") locPath = [];
  advanceFrom(stage);
}

/** Go to the next slot that still needs filling, or to "done". */
function advanceFrom(key) {
  const start = STAGES.findIndex((s) => s.key === key);
  for (let step = 1; step <= STAGES.length; step++) {
    const s = STAGES[(start + step) % STAGES.length].key;
    if (!sel[s] && !skipped[s] && applicable(s)) {
      stage = s;
      render();
      $("logInput").focus();
      return;
    }
  }
  stage = "done";
  render();
  $("logInput").focus();
}

function goToStage(key) {
  stage = key;
  if (key === "location") locPath = [];
  render();
}

function resetEntry() {
  sel = emptySelection();
  skipped = {};
  locPath = [];
  stage = "callsign";
  $("logInput").value = "";
  render();
}

function clearAll() {
  if (busy) return;
  resetEntry();
  setStatus("Cleared.", "ok");
  $("logInput").focus();
}

/* ==========================================================================
 * Logging
 * ========================================================================*/

function logEntry() {
  if (busy) return;
  const notes = $("logInput").value.trim();
  const anyPicked = STAGES.some((s) => sel[s.key]);

  if (!anyPicked) {
    if (notes === "") { setStatus("Pick some buttons or type shorthand first.", "err"); return; }
    pendingRecent = null;
    send({ kind: "text", text: notes });
    return;
  }

  const details = [sel.incident ? sel.incident.label : "", notes].filter(Boolean).join(" - ");
  const fields = [
    sel.callsign ? sel.callsign.label : "",
    sel.department ? sel.department.label : "",
    sel.location ? sel.location.label : "",
    details,
  ];
  pendingRecent = { callsign: sel.callsign, department: sel.department, incident: sel.incident, location: sel.location };
  send({ kind: "entry", fields });
}

function send(obj) {
  busy = true;
  $("logInput").disabled = true;
  $("logBtn").disabled = true;
  setStatus("Logging…", "busy");
  transport.send(obj);
  busyTimer = setTimeout(() => {
    busy = false;
    $("logInput").disabled = false;
    $("logBtn").disabled = false;
    setStatus("Error: no reply from Excel. Check the Log sheet before logging again.", "err");
  }, SEND_TIMEOUT_MS);
}

/* ==========================================================================
 * Recent entries (stored in this browser only)
 * ========================================================================*/

function loadRecents() {
  try { return JSON.parse(localStorage.getItem(RECENTS_KEY) || "[]"); } catch (e) { return []; }
}

function addRecent(entry) {
  const list = loadRecents().filter((r) => recentText(r) !== recentText(entry));
  list.unshift(entry);
  try { localStorage.setItem(RECENTS_KEY, JSON.stringify(list.slice(0, RECENTS_MAX))); } catch (e) { /* storage off */ }
  renderRecents();
}

function recentText(r) {
  const parts = [];
  if (r.callsign) parts.push(r.callsign.code ? r.callsign.code.toUpperCase() : r.callsign.label);
  if (r.incident) parts.push(r.incident.label);
  else if (r.department) parts.push(r.department.label);
  if (r.location) parts.push(r.location.label);
  return parts.join(" › ");
}

function useRecent(r) {
  if (busy) return;
  sel = emptySelection();
  skipped = {};
  for (const s of STAGES) {
    const saved = r[s.key];
    sel[s.key] = saved ? (menu[s.key].find((i) => i.label === saved.label) || saved) : null;
    if (!saved) skipped[s.key] = true;
  }
  stage = "done";
  render();
  setStatus("Loaded a recent entry. Add notes, then Log entry.", "ok");
  $("logInput").focus();
}

function renderRecents() {
  const box = $("recents");
  box.innerHTML = "";
  const list = loadRecents();
  if (list.length === 0) return;
  const label = document.createElement("span");
  label.className = "label";
  label.textContent = "Recent";
  box.appendChild(label);
  for (const r of list) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip";
    chip.textContent = recentText(r);
    chip.title = recentText(r);
    const c = (r.incident && r.incident.colour) || (r.department && r.department.colour);
    if (c) chip.style.setProperty("--chip", c);
    chip.addEventListener("click", () => useRecent(r));
    box.appendChild(chip);
  }
}

/* ==========================================================================
 * Rendering
 * ========================================================================*/

function render() {
  renderDocket();
  renderStage();
  const anyPicked = STAGES.some((s) => sel[s.key]);
  $("logInput").placeholder = anyPicked
    ? "Notes (optional), then Enter"
    : "Pick buttons above, or type shorthand: l1s ps ed14 ws";
}

function renderDocket() {
  const docket = $("docket");
  docket.innerHTML = "";
  for (const s of STAGES) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "slot" + (stage === s.key ? " active" : "");
    const item = sel[s.key];
    let text;
    let empty = false;
    if (item) text = item.label;
    else if (skipped[s.key]) { text = "Skipped"; empty = true; }
    else if (s.key === "incident" && sel.department && !applicable("incident")) { text = "No list"; empty = true; }
    else { text = "—"; empty = true; }

    b.innerHTML = '<span class="slot-name"></span><span class="slot-value"></span>';
    b.querySelector(".slot-name").textContent = s.name;
    const v = b.querySelector(".slot-value");
    v.textContent = text;
    v.title = text;
    if (empty) v.classList.add("empty");
    if (item && item.colour) b.style.setProperty("--slot-colour", item.colour);
    b.addEventListener("click", () => goToStage(s.key));
    docket.appendChild(b);
  }
}

function renderStage() {
  const grid = $("grid");
  const title = $("stageTitle");
  grid.innerHTML = "";
  grid.classList.remove("compact");
  $("backBtn").hidden = !(stage === "location" && locPath.length > 0);
  $("skipBtn").hidden = stage === "done";

  if (!menuLoaded) {
    title.textContent = "Loading…";
    return;
  }

  if (stage === "done") {
    title.textContent = "Ready to log";
    notice(grid, "Add notes below if you need to, then press Enter or Log entry. Click any slot above to change it.");
    return;
  }

  const def = STAGES.find((s) => s.key === stage);
  title.textContent = def.prompt;
  if (stage === "location" && locPath.length) {
    const path = document.createElement("span");
    path.className = "path";
    path.textContent = "  " + locPath.join(" › ");
    title.appendChild(path);
  }

  let items = [];
  let folders = [];
  if (stage === "incident") items = incidentsFor(sel.department);
  else if (stage === "location") {
    const node = currentLocNode();
    folders = node.children;
    items = node.items;
  } else items = menu[stage];

  if (items.length === 0 && folders.length === 0) {
    notice(grid, stage === "incident"
      ? "This department has no incident buttons. Click Skip, or add Incident rows on Inputs with this department as the group."
      : "No buttons for this slot yet. Add rows on the Inputs sheet with this Button type, then click Reload buttons.");
    return;
  }

  for (const f of folders) {
    // A group holding just one location picks it straight away.
    const only = f.children.length === 0 && f.items.length === 1 ? f.items[0] : null;
    const t = tile(f.name, "", f.colour, () => {
      if (only) pick("location", only);
      else { locPath.push(f.name); render(); }
    });
    t.classList.add("folder");
    const count = document.createElement("span");
    count.className = "count";
    count.textContent = only ? "" : countItems(f) + " ›";
    t.appendChild(count);
    grid.appendChild(t);
  }

  const labels = shortLabels(items);
  if (folders.length === 0 && items.length > 0 && labels.every((l) => l.length <= 4)) grid.classList.add("compact");

  items.forEach((item, idx) => {
    const t = tile(labels[idx], grid.classList.contains("compact") ? "" : item.code, item.colour, () => pick(stage, item));
    t.title = item.label + (item.code ? "  (" + item.code + ")" : "");
    if (sel[stage] === item) t.classList.add("chosen");
    grid.appendChild(t);
  });
}

function tile(label, code, colour, onClick) {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "tile";
  const text = document.createElement("span");
  text.textContent = label;
  b.appendChild(text);
  if (code) {
    const c = document.createElement("span");
    c.className = "code";
    c.textContent = code;
    b.appendChild(c);
  }
  const bg = colour || "#475569";
  b.style.setProperty("--tile", bg);
  b.style.setProperty("--tile-ink", inkFor(bg));
  b.addEventListener("click", onClick);
  return b;
}

function notice(grid, text) {
  const p = document.createElement("div");
  p.className = "notice";
  p.textContent = text;
  grid.appendChild(p);
}

/**
 * When a set of siblings all start with the same words ("External Door 1",
 * "External Door 2"…), show just what differs ("1", "2"…). The full label
 * is still what gets logged.
 */
function shortLabels(items) {
  const labels = items.map((i) => i.label);
  if (items.length < 4) return labels;
  const words = labels.map((l) => l.split(/\s+/));
  let p = 0;
  while (words.every((w) => w.length > p + 1 && w[p].toLowerCase() === words[0][p].toLowerCase())) p++;
  return p === 0 ? labels : words.map((w) => w.slice(p).join(" "));
}

/** White or near-black text, whichever reads better on the tile colour. */
function inkFor(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return "#fff";
  const n = parseInt(m[1], 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => {
    v /= 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  });
  const lum = 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
  return lum > 0.22 ? "#0f172a" : "#ffffff";
}

function setStatus(message, kind) {
  const el = $("status");
  el.textContent = message;
  el.className = kind === "ok" ? "status-ok" : kind === "err" ? "status-err" : "status-busy";
}
