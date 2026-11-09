// Set this to match where you host these files, e.g. "https://yourname.github.io/venue-logger-dialog"
const BASE_URL = "https://calvinwillans.github.io/loggerapp";
const DIALOG_URL = BASE_URL + "/dialog.html";

const INPUTS_SHEET = "Inputs";
// Inputs columns used to build the dialog's buttons (A = code, B = label).
// Columns are found by their header text in row 1 where possible, so
// inserting a column on Inputs can't knock the mapping out of line. The
// numbers below are only the fallback when a header isn't found.
const INPUTS_TYPE_COL = 2;   // C - Callsign / Department / Incident / Location
const INPUTS_GROUP_COL = 3;  // D - group path, e.g. "Blocks/200s" or "Level 2;Toilets"
const INPUTS_COLOUR_COL = 4; // E - hex colour, e.g. #1D4ED8

// Header text this add-in looks for, per column. Matching ignores case,
// spaces and punctuation, so "Button Row", "button-row" and "ButtonRow" are
// all the same thing.
const INPUTS_HEADERS = {
  code: ["code", "shorthand"],
  label: ["label", "expansion"],
  type: ["buttontype", "type"],
  group: ["buttongroup", "group", "folder"],
  colour: ["colour", "color"],
  home: ["homelocation", "home", "defaultlocation"],
  row: ["buttonrow", "row"],
  parentRow: ["parentrow", "folderrow", "parentbuttonrow"],
};
const INPUTS_HOME_COL = 5;   // F - callsigns only: a location label, or a location folder path
const INPUTS_ROW_COL = 6;    // G - row number for this button on its own screen
const INPUTS_PARENT_ROW_COL = 7; // H - row number for the folder button this one sits in
const BUTTON_TYPES = ["callsign", "department", "incident", "location"];

// Dialog size, as a percentage of the screen.
const DIALOG_WIDTH = 50;
const DIALOG_HEIGHT = 75;
// Office caps each message to the dialog, so the button data goes in pieces.
const MESSAGE_CHUNK_CHARS = 15000;
const LOG_SHEET = "Log";
const LOG_FIRST_DATA_ROW = 8; // first entry always lands here, matching the template
const LOG_MAX_ROW = 5000; // generous ceiling when scanning for the next empty row

// Last column of the log table itself (Time .. Secondary Action). When a
// backfilled milestone needs a row pushed in, only A:G shifts down - the
// Live Counts panel to the right (I2:K15) stays where it is. Inserting a
// whole sheet row used to split that panel in half.
const LOG_LAST_COLUMN = "G";

const ACTIONS_SHEET = "Actions";
const DEPARTMENT_COLUMN = "C";
const ACTION_COLUMN = "F";

// Column G: a second dropdown, for incidents that need more than one action
// (WAPOL and medics, say). Its list lives in a column on the Actions sheet
// under this header. The add-in only fills that column in when it's missing
// or empty, seeding it with every action from every department - after that
// it's yours to edit, and nothing here overwrites it.
const SECONDARY_ACTION_COLUMN = "G";
const ALL_ACTIONS_HEADER = "All Actions";

// An event runs past midnight, so a time earlier than this counts as "after
// midnight" and sorts to the end of the Log rather than the beginning.
// 6am suits a venue whose events start in the afternoon or evening.
const DAY_ROLLOVER_HOUR = 6;

// What to do when a milestone cell has something in it that isn't a time
// (a word, a header, a stray note). false = skip it and say so in the sweep
// report. true = log it anyway, stamped with the current clock time.
// This was effectively true before, which is how the "Time Confirmed"
// header in C104 got logged as a whitelevel check on every fresh copy.
const LOG_NON_TIME_WITH_CLOCK = false;

const MILESTONE_SHEET = "Event Checks";
const STATE_SHEET = "_LoggerState";

// Older builds kept a record of logged milestones in the workbook's add-in
// settings. That record travelled with every copy of the file and couldn't
// be shared reliably between people co-authoring, so it is no longer used:
// the Log itself is the record. Reset Milestones deletes the old setting.
const OLD_MILESTONE_STATE_KEY = "VenueLogger.milestoneState";

// Cell on "Event Checks" -> the row to put in the Log when someone types a
// time into that cell. Add more entries here in the same shape as they come
// up - leave any column "" to leave it blank on the logged row.
//
// Every address here must be the time cell on that check's own row. Row 104
// is the "White Level Checks" section header, so the whitelevel entries start
// at row 105, and Venue Clear is row 110.
//
// alsoOn: for a cell filled by FORMULA from another cell, the cell people
// actually type into. E50 is =IF(F14="","",F14), so the milestone fires when
// someone types the WAPOL Onsite time into F14.
const MILESTONE_MAP = {
  F23: { b: "Base", c: "All Call", d: "", e: "All departments switch over to event channels" },
  F25: { b: "Base", c: "", d: "", e: "Fire panel in event mode. EWIS switched to manual" },
  E50: { b: "Base", c: "", d: "", e: "WAPOL arrived onsite", alsoOn: ["F14"] },
  E52: { b: "Base", c: "All Call", d: "", e: "RAC Local Lounge now open" },
  E62: { b: "Base", c: "All Call", d: "", e: "VIP Doors Open" },
  E63: { b: "Base", c: "All Call", d: "", e: "External doors now open" },
  E65: { b: "Base", c: "All Call", d: "", e: "Internal doors now open" },
  C73: { b: "Base", c: "All Call", d: "", e: "Start of support act" },
  D73: { b: "Base", c: "All Call", d: "", e: "End of support act" },
  C74: { b: "Base", c: "All Call", d: "", e: "Start of main act/game" },
  C75: { b: "Base", c: "All Call", d: "", e: "Start of Intermission/Halftime" },
  D75: { b: "Base", c: "All Call", d: "", e: "End of Intermission/Halftime" },
  C87: { b: "Base", c: "All Call", d: "", e: "End of show/game. Prepare for egress" },
  C105: { b: "Security", c: "Base", d: "", e: "Level 4 whitelevel checks completed" },
  C106: { b: "Security", c: "Base", d: "", e: "Level 3 whitelevel checks completed" },
  C107: { b: "Security", c: "Base", d: "", e: "Level 2 whitelevel checks completed" },
  C108: { b: "Security", c: "Base", d: "", e: "Level 1 whitelevel checks completed" },
  C109: { b: "Security", c: "Base", d: "", e: "Ground floor whitelevel checks completed" },
  C110: { b: "Base", c: "All Call", d: "", e: "Venue clear. Switch to non-event channels" },
  // Row 119 = "Handover of Channel 1 back to CRO and Base sign off", the last
  // step of the night. (Row 115 is "External Checks Started", which is why
  // this used to log as soon as external checks began.)
  C119: { b: "Security", c: "Base", d: "", e: "External checks complete. Channel 1 handover to CRO & EWIS back to auto" },
};

let dialog = null;

// Runs once, as soon as this page loads. With the shared runtime + a "long"
// lifetime in the manifest, that's now on workbook open rather than only
// when the ribbon button is clicked or a dialog message arrives.
let watchersRegistered = false;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // Office can run onReady more than once in a session. Registering the
    // same onChanged handler twice means every edit sweeps twice.
    if (watchersRegistered) return;
    watchersRegistered = true;
    try {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    } catch (err) {
      console.error("setStartupBehavior failed:", err);
    }
    await registerDepartmentWatcher();
    await registerMilestoneWatcher();
    // No milestone sweep here. With several people opening the workbook,
    // each one sweeping on open could log the same milestone; milestones log
    // only from the edit that filled them in.
    await refreshSecondaryOnOpen();
  }
});

/** Puts the column G dropdown on rows already in the Log when the file opens. */
async function refreshSecondaryOnOpen() {
  try {
    await Excel.run(async (context) => {
      const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
      logSheet.load("isNullObject");
      await context.sync();
      if (logSheet.isNullObject) return;
      await refreshSecondaryValidation(context, logSheet);
    });
  } catch (err) {
    console.error("refreshSecondaryValidation on open failed:", err);
  }
}

/**
 * Registers a listener on the Log sheet: any edit to column C (Department)
 * refreshes that row's Action Taken (F) dropdown to match the corresponding
 * column on the Actions sheet. Fully open-ended - add a new column to
 * Actions with any header text and it's usable immediately, no other
 * changes needed.
 */
async function registerDepartmentWatcher() {
  await Excel.run(async (context) => {
    const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
    logSheet.load("isNullObject");
    await context.sync();

    if (logSheet.isNullObject) {
      console.error(`"${LOG_SHEET}" sheet not found - department watcher not registered.`);
      return;
    }

    logSheet.onChanged.add(onLogSheetChanged);
    await context.sync();
    console.log("Department watcher registered.");
  });
}

async function onLogSheetChanged(eventArgs) {
  if (!isLocalEdit(eventArgs)) return; // the copy that made the edit handles it
  await Excel.run(async (context) => {
    const logSheet = context.workbook.worksheets.getItem(LOG_SHEET);

    const changedRange = logSheet.getRange(eventArgs.address.split("!").pop());
    changedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
    await context.sync();

    const firstDataRowIndex = LOG_FIRST_DATA_ROW - 1;
    const endRow = changedRange.rowIndex + changedRange.rowCount - 1;
    if (endRow < firstDataRowIndex) return;

    // Column G follows whether the row has anything in it, whatever changed.
    await refreshSecondaryValidation(context, logSheet);

    const deptColIndex = columnLetterToIndex(DEPARTMENT_COLUMN);
    const touchesDept =
      changedRange.columnIndex <= deptColIndex &&
      deptColIndex < changedRange.columnIndex + changedRange.columnCount;
    if (!touchesDept) return;

    const startRow = Math.max(changedRange.rowIndex, firstDataRowIndex);

    const sources = await buildActionSources(context);

    for (let r = startRow; r <= endRow; r++) {
      const rowNum = r + 1; // 1-based
      const deptCell = logSheet.getRange(`${DEPARTMENT_COLUMN}${rowNum}`);
      deptCell.load("values");
      await context.sync();

      const department = (deptCell.values[0][0] || "").toString().trim().toLowerCase();
      const actionCell = logSheet.getRange(`${ACTION_COLUMN}${rowNum}`);
      actionCell.dataValidation.clear();

      const source = sources[department];
      if (source) {
        actionCell.dataValidation.rule = { list: { inCellDropDown: true, source: source } };
      }
      await context.sync();
    }
  });
}

/** Reads Actions!row1 headers and returns { lowercased-header -> "=Actions!$X$2:$X$N" }. */
async function buildActionSources(context) {
  const actionsSheet = context.workbook.worksheets.getItemOrNullObject(ACTIONS_SHEET);
  actionsSheet.load("isNullObject");
  await context.sync();
  if (actionsSheet.isNullObject) return {};

  const used = actionsSheet.getUsedRange();
  used.load("values, rowIndex, columnIndex");
  await context.sync();

  const sources = {};
  const values = used.values;
  const headerRow = values[0];
  const originRow = used.rowIndex;
  const originCol = used.columnIndex;

  for (let c = 0; c < headerRow.length; c++) {
    const header = (headerRow[c] || "").toString().trim();
    if (!header) continue;

    let lastDataIndex = 0;
    for (let r = 1; r < values.length; r++) {
      if ((values[r][c] || "").toString().trim() !== "") lastDataIndex = r;
    }
    if (lastDataIndex === 0) continue; // header with nothing under it yet

    const colLetter = columnIndexToLetter(originCol + c);
    const firstDataRow = originRow + 2;               // 1-based, just below header
    const lastDataRow = originRow + lastDataIndex + 1; // 1-based, last item

    sources[header.toLowerCase()] = `=${ACTIONS_SHEET}!$${colLetter}$${firstDataRow}:$${colLetter}$${lastDataRow}`;
  }
  return sources;
}

/**
 * Returns the data validation source for the secondary-action list, e.g.
 * "=Actions!$J$2:$J$47".
 *
 * If the "All Actions" column already has entries, it is used exactly as it
 * stands - trim it, reorder it, add your own wording, and it stays that way.
 * Only when the column is missing or empty is it seeded with every action
 * from every department. To start again from the full list, clear the column
 * (or delete it) and reopen the workbook.
 */
async function buildAllActionsSource(context) {
  const actionsSheet = context.workbook.worksheets.getItemOrNullObject(ACTIONS_SHEET);
  actionsSheet.load("isNullObject");
  await context.sync();
  if (actionsSheet.isNullObject) return "";

  const used = actionsSheet.getUsedRange();
  used.load("values, rowIndex, columnIndex");
  await context.sync();

  const values = used.values;
  if (!values || values.length === 0) return "";
  const headerRow = values[0];
  const originRow = used.rowIndex;
  const originCol = used.columnIndex;

  // Find the list column: an existing "All Actions" header, or the first free
  // column to the right of everything else.
  let targetIndex = -1;
  for (let c = 0; c < headerRow.length; c++) {
    if ((headerRow[c] || "").toString().trim().toLowerCase() === ALL_ACTIONS_HEADER.toLowerCase()) targetIndex = c;
  }
  const isNew = targetIndex === -1;
  if (isNew) targetIndex = headerRow.length;

  const sourceFor = (count) => {
    const colLetter = columnIndexToLetter(originCol + targetIndex);
    return `=${ACTIONS_SHEET}!$${colLetter}$${originRow + 2}:$${colLetter}$${originRow + count + 1}`;
  };

  // Already has a list? Leave it completely alone.
  if (!isNew) {
    let lastIndex = 0;
    for (let r = 1; r < values.length; r++) {
      if ((values[r][targetIndex] || "").toString().trim() !== "") lastIndex = r;
    }
    if (lastIndex > 0) return sourceFor(lastIndex);
  }

  // Empty or missing: seed it once with every action from every department.
  const seen = new Set();
  const all = [];
  for (let c = 0; c < headerRow.length; c++) {
    if (c === targetIndex) continue;
    for (let r = 1; r < values.length; r++) {
      const text = (values[r][c] || "").toString().trim();
      if (text === "") continue;
      const key = text.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      all.push(text);
    }
  }
  if (all.length === 0) return "";

  const rows = [[ALL_ACTIONS_HEADER]].concat(all.map((a) => [a]));
  actionsSheet.getRangeByIndexes(originRow, originCol + targetIndex, rows.length, 1).values = rows;
  await context.sync();
  console.log(`Seeded "${ALL_ACTIONS_HEADER}" with ${all.length} actions - edit that column to change the column G dropdown.`);

  return sourceFor(all.length);
}

/**
 * Puts the secondary-action dropdown on every populated Log row and takes it
 * off the empty ones below. Log rows are always filled top down, so the
 * populated block runs from LOG_FIRST_DATA_ROW to the last used row.
 */
async function refreshSecondaryValidation(context, logSheet) {
  const source = await buildAllActionsSource(context);
  if (!source) return;

  const nextRow = await findNextLogRow(context, logSheet);
  const lastRow = nextRow - 1;

  if (lastRow >= LOG_FIRST_DATA_ROW) {
    const filled = logSheet.getRange(`${SECONDARY_ACTION_COLUMN}${LOG_FIRST_DATA_ROW}:${SECONDARY_ACTION_COLUMN}${lastRow}`);
    filled.dataValidation.clear();
    await context.sync();
    filled.dataValidation.rule = { list: { inCellDropDown: true, source: source } };
  }

  if (nextRow <= LOG_MAX_ROW) {
    logSheet
      .getRange(`${SECONDARY_ACTION_COLUMN}${nextRow}:${SECONDARY_ACTION_COLUMN}${LOG_MAX_ROW}`)
      .dataValidation.clear();
  }
  await context.sync();
}

function columnLetterToIndex(letters) {
  let index = 0;
  for (const ch of letters.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
  return index - 1;
}

function columnIndexToLetter(index) {
  let letters = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/**
 * Registers a listener on "Event Checks". Any edit there triggers a sweep of
 * every mapped cell, so it doesn't matter which cell was touched.
 */
async function registerMilestoneWatcher() {
  await Excel.run(async (context) => {
    const checksSheet = context.workbook.worksheets.getItemOrNullObject(MILESTONE_SHEET);
    checksSheet.load("isNullObject");
    await context.sync();

    if (checksSheet.isNullObject) {
      console.error(`"${MILESTONE_SHEET}" sheet not found - milestone watcher not registered.`);
      return;
    }

    checksSheet.onChanged.add(onMilestoneSheetChanged);
    await context.sync();
    console.log("Milestone watcher registered.");
  });
}

/**
 * True for an edit made in THIS copy of Excel. The workbook is open on
 * several PCs at once, and every open copy hears about every edit. Only the
 * copy where the edit was actually made reacts to it - otherwise each copy
 * logs the same milestone, which is what caused the doubled and tripled rows.
 */
function isLocalEdit(eventArgs) {
  const source = eventArgs && eventArgs.source;
  return !source || source.toString().toLowerCase() === "local";
}

async function onMilestoneSheetChanged(eventArgs) {
  if (!isLocalEdit(eventArgs)) return;
  const addresses = milestonesTouchedBy(eventArgs && eventArgs.address);
  if (addresses.length === 0) return;
  const result = await sweepMilestones(addresses);
  if (result.error) console.error("Milestone sweep failed:", result.error);
  for (const line of result.report || []) console.log("  " + line);
}

/** "C73" -> { col: 2, row: 73 } (0-based column, 1-based row). */
function parseCell(a1) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(a1.trim());
  if (!m) return null;
  return { col: columnLetterToIndex(m[1].toUpperCase()), row: parseInt(m[2], 10) };
}

/**
 * The milestones an edit to this address should check: every mapped cell
 * inside the edited range, plus any mapped cell that copies an edited cell
 * by formula (listed in its alsoOn). Handles single cells, pasted blocks
 * ("C105:C109") and multi-area addresses ("C73,D75").
 */
function milestonesTouchedBy(address) {
  if (!address) return [];
  const areas = address.toString().split(",").map((part) => {
    const bare = part.split("!").pop();
    const ends = bare.split(":");
    const a = parseCell(ends[0]);
    const b = parseCell(ends[1] || ends[0]);
    if (!a || !b) return null;
    return {
      c1: Math.min(a.col, b.col), c2: Math.max(a.col, b.col),
      r1: Math.min(a.row, b.row), r2: Math.max(a.row, b.row),
    };
  }).filter(Boolean);

  const inside = (cellAddress) => {
    const c = parseCell(cellAddress);
    return !!c && areas.some((ar) => c.col >= ar.c1 && c.col <= ar.c2 && c.row >= ar.r1 && c.row <= ar.r2);
  };

  return Object.keys(MILESTONE_MAP).filter((address) => {
    if (inside(address)) return true;
    return (MILESTONE_MAP[address].alsoOn || []).some(inside);
  });
}

/**
 * One comparable form for a Details string: non-breaking spaces become
 * ordinary ones, runs of whitespace collapse, case is ignored. Without this,
 * a double space or a stray nbsp makes a row look "not yet logged".
 */
function normaliseDetails(value) {
  return (value === null || value === undefined ? "" : value.toString())
    .replace(/ /g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/**
 * Checks the given milestone cells (all of them if none are given) and logs
 * any that need it. Runs in the write queue, so it never overlaps another
 * write from this PC.
 */
function sweepMilestones(addresses) {
  return withLogLock(() => runMilestoneSweep(addresses || Object.keys(MILESTONE_MAP)));
}

// Every write to the Log from this PC - logger entries and milestones - goes
// through this queue, one at a time. Placement reads the Log, decides on a
// row, then inserts; if two writes interleaved, the second would be working
// from a row number the first had just shifted.
let logWriteQueue = Promise.resolve();

function withLogLock(task) {
  const run = logWriteQueue.then(() => task(), () => task());
  logWriteQueue = run.catch(() => {});
  return run;
}

/**
 * Works out where a row with this time belongs in the Log: just above the
 * first entry with a later time, or below the last entry if nothing is
 * later. Rows with the same time keep the order they arrived in.
 *
 * Returns { row, reason }; row > LOG_MAX_ROW means the Log is full.
 */
async function planLogPlacement(context, logSheet, seconds) {
  const range = logSheet.getRange(`A${LOG_FIRST_DATA_ROW}:A${LOG_MAX_ROW}`);
  range.load("values, text");
  await context.sync();

  const values = range.values;
  const texts = range.text;
  const key = timeSortKey(seconds);

  let blankRow = LOG_MAX_ROW + 1;
  let candidate = null; // the first row with a later time

  for (let i = 0; i < values.length; i++) {
    const rowNumber = LOG_FIRST_DATA_ROW + i;
    const held = values[i][0];
    const shown = texts[i][0];
    const blank =
      (held === null || held === undefined || held.toString().trim() === "") &&
      (shown === null || shown === undefined || shown.toString().trim() === "");

    if (blank) { blankRow = rowNumber; break; } // reached the bottom

    if (candidate !== null) continue; // already know where it goes, just need the bottom

    // Prefer the stored value, fall back to what's displayed.
    let existing = parseTimeToSeconds(held);
    if (existing === null) existing = parseTimeToSeconds(shown);
    if (existing === null) continue; // not a time - leave that row alone

    if (timeSortKey(existing) > key) {
      candidate = { row: rowNumber, existingShown: (shown || held || "").toString().trim() };
    }
  }

  // The Log has to be able to grow by one row.
  if (blankRow > LOG_MAX_ROW) return { row: LOG_MAX_ROW + 1, reason: "the Log is full" };

  if (candidate === null) return { row: blankRow, reason: "added at the bottom" };
  return { row: candidate.row, reason: `inserted above ${candidate.existingShown} in row ${candidate.row}` };
}

/**
 * Puts one row into the Log at the place its time belongs. fields =
 * [callsign, department, location, details]. Returns { row, reason }; row 0
 * means the Log is full. Call only from inside withLogLock.
 *
 * A new row is ALWAYS pushed in (A:G only, so the Live Counts panel beside
 * the Log isn't touched), even at the bottom, and the push and the write go
 * to Excel together. With several people logging at once, writing into "the
 * first blank row" lets two of them pick the same row and one entry
 * overwrites the other. Two inserts at the same spot become two rows.
 */
async function writeLogRow(context, logSheet, seconds, fields) {
  const plan = await planLogPlacement(context, logSheet, seconds);
  if (plan.row > LOG_MAX_ROW) return { row: 0, reason: "the Log is full" };

  const f = [0, 1, 2, 3].map((i) => (fields[i] === null || fields[i] === undefined ? "" : fields[i].toString().trim()));
  const rowValues = [[secondsToClock(seconds), f[0], f[1], f[2], f[3]]];

  let row = plan.row;
  try {
    logSheet.getRange(`A${row}:${LOG_LAST_COLUMN}${row}`).insert(Excel.InsertShiftDirection.down);
    logSheet.getRange(`A${row}:E${row}`).values = rowValues;
    await context.sync();
  } catch (err) {
    // Inserting can be refused (a protected sheet, say). Rather than lose
    // the entry, add it in the first blank row instead.
    console.error("Could not insert a row in the Log:", err);
    row = await findNextLogRow(context, logSheet);
    if (row > LOG_MAX_ROW) return { row: 0, reason: "the Log is full" };
    logSheet.getRange(`A${row}:E${row}`).values = rowValues;
    await context.sync();
    return { row: row, reason: "could not insert a row, added at the bottom instead" };
  }
  return { row: row, reason: plan.reason };
}

/**
 * The time a dialog entry should carry: when the operator started it (first
 * button press, or first character typed), sent by the dialog as startedAt.
 * Falls back to now if the dialog didn't send one (an older cached dialog).
 */
function entryStartSeconds(startedAt) {
  const parsed = startedAt === undefined || startedAt === null || startedAt === "" ? null : parseTimeToSeconds(startedAt);
  return parsed === null ? { seconds: nowSeconds(), fromStart: false } : { seconds: parsed, fromStart: true };
}

// Milestones the Reset button has re-armed on this PC: each may log once
// more, even though its wording is already in the Log.
const rearmed = new Set();

/**
 * Logs the given milestone cells, each at the time typed into it.
 *
 * The Log itself is the shared record of what has been logged, because it
 * is the one thing every open copy of the workbook sees: a milestone whose
 * Details wording is already in the Log is skipped, unless Reset Milestones
 * re-armed it. Nothing is kept in hidden state, so nothing carries over when
 * the workbook is copied for the next event.
 */
async function runMilestoneSweep(addresses) {
  try {
    return await Excel.run(async (context) => {
      const checksSheet = context.workbook.worksheets.getItemOrNullObject(MILESTONE_SHEET);
      const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
      checksSheet.load("isNullObject");
      logSheet.load("isNullObject");
      await context.sync();

      if (checksSheet.isNullObject) return { checked: 0, logged: 0, error: `"${MILESTONE_SHEET}" sheet not found.`, report: [] };
      if (logSheet.isNullObject) return { checked: 0, logged: 0, error: `"${LOG_SHEET}" sheet not found.`, report: [] };

      const cells = addresses.map((a) => {
        const r = checksSheet.getRange(a);
        r.load("text, values");  // text for a real time, values as the backstop
        return r;
      });

      // What's already in the Log. VALUES, not text: text is what the cell
      // displays, and Excel clips that to the column width.
      const detailsRange = logSheet.getRange(`E${LOG_FIRST_DATA_ROW}:E${LOG_MAX_ROW}`);
      detailsRange.load("values");
      await context.sync();

      const already = new Set();
      for (const row of detailsRange.values) {
        const text = normaliseDetails(row[0]);
        if (text !== "") already.add(text);
      }

      const report = [];
      const toLog = [];
      addresses.forEach((address, i) => {
        const shownRaw = cells[i].text[0][0];
        const heldRaw = cells[i].values[0][0];
        const shown = (shownRaw === null || shownRaw === undefined ? "" : shownRaw).toString().trim();
        const held = (heldRaw === null || heldRaw === undefined ? "" : heldRaw).toString().trim();
        const seen = shown !== "" ? shown : held;

        if (seen === "" || isNotApplicable(shown) || isNotApplicable(held)) {
          report.push(`${address}: empty, skipped`);
          return;
        }

        const mapping = MILESTONE_MAP[address];
        const key = normaliseDetails(mapping.e);
        if (key === "") return;

        if (already.has(key) && !rearmed.has(address)) {
          report.push(`${address}: "${seen}" already in the Log, skipped`);
          return;
        }

        // The time comes from the cell itself, not the clock, so a milestone
        // backfilled an hour later still carries the time it happened.
        let seconds = parseTimeToSeconds(heldRaw);
        if (seconds === null) seconds = parseTimeToSeconds(shown);
        const parsed = seconds !== null;
        if (!parsed) {
          if (!LOG_NON_TIME_WITH_CLOCK) {
            report.push(`${address}: "${seen}" is not a time, skipped`);
            return;
          }
          seconds = nowSeconds();
        }

        toLog.push({ address: address, mapping: mapping, seconds: seconds, parsed: parsed, seen: seen });
        already.add(key);
      });

      let logged = 0;
      for (const entry of toLog) {
        const m = entry.mapping;
        const placed = await writeLogRow(context, logSheet, entry.seconds, [m.b, m.c, m.d, m.e]);
        if (placed.row === 0) {
          report.push(`${entry.address}: ${placed.reason}`);
          continue;
        }
        rearmed.delete(entry.address);
        report.push(
          `${entry.address}: "${entry.seen}" -> row ${placed.row} at ${secondsToClock(entry.seconds)}` +
            (entry.parsed ? "" : " (no time found in the cell, used the clock)") +
            ` - ${placed.reason}`
        );
        logged++;
      }

      if (logged > 0) {
        // The dropdown refresh must never undo a successful log: if Excel
        // refuses the validation change, the rows still stand.
        try {
          await refreshSecondaryValidation(context, logSheet);
        } catch (err) {
          console.error("Secondary action dropdown after milestone sweep:", err);
        }
      }
      return { checked: addresses.length, logged: logged, error: null, report: report };
    });
  } catch (err) {
    console.error("sweepMilestones failed:", err);
    return { checked: 0, logged: 0, error: err.message || String(err), report: [] };
  }
}

/** Wording for the dialog's status line. */
function sweepSummary(result) {
  if (result.error) return "Milestone check failed: " + result.error;
  if (result.logged === 0) return `${result.checked} milestones checked, none new.`;
  return `${result.checked} milestones checked, ${result.logged} logged.`;
}

/**
 * Ribbon-ready (not on the ribbon yet): check EVERY milestone cell now and
 * log any filled-in one whose wording isn't in the Log - a deliberate
 * catch-up for times typed in by someone who didn't have the add-in running.
 * Re-typing that one time does the same job without a button.
 */
async function syncMilestones(event) {
  const result = await sweepMilestones();
  console.log(sweepSummary(result));
  // Full per-cell reasoning, so a milestone that behaves oddly can be traced.
  for (const line of result.report || []) console.log("  " + line);
  sendToDialog({ kind: "status", ok: !result.error, text: sweepSummary(result) });
  if (event) event.completed();
}

/** Same "next empty row" scan handleLogEntry uses, kept separate so this path can never affect the dialog's own logging. */
async function findNextLogRow(context, logSheet) {
  const scanRange = logSheet.getRange(`A${LOG_FIRST_DATA_ROW}:A${LOG_MAX_ROW}`);
  scanRange.load("text");
  await context.sync();

  const colText = scanRange.text;
  for (let i = 0; i < colText.length; i++) {
    if ((colText[i][0] || "").toString().trim() === "") return LOG_FIRST_DATA_ROW + i;
  }
  return LOG_MAX_ROW + 1;
}

/**
 * Reads a time out of a milestone cell and returns seconds since midnight,
 * or null if there is no time to be had.
 *
 * Handles what people actually type, plus what Excel stores underneath:
 *   a real time value (0.8125, i.e. a fraction of a day, or a date serial)
 *   "19:30"  "19:30:45"  "7.30pm"  "7:30 PM"  "1930"
 */
function parseTimeToSeconds(raw) {
  if (raw === null || raw === undefined) return null;

  // A real Excel time is a number: whole part = date, fraction = time of day.
  if (typeof raw === "number" && isFinite(raw)) {
    const fraction = raw - Math.floor(raw);
    if (fraction > 0) return Math.round(fraction * 86400);
    // A whole number isn't a time value, but people do type 1930 into a cell
    // and Excel stores that as the number 1930.
    if (raw >= 1 && raw <= 2359) {
      const hours = Math.floor(raw / 100);
      const minutes = raw % 100;
      if (hours <= 23 && minutes <= 59) return hours * 3600 + minutes * 60;
    }
    return null;
  }

  const text = raw.toString().trim();
  if (text === "") return null;

  // 19:30, 19:30:45, 7.30 pm
  let m = text.match(/^(\d{1,2})[:.](\d{1,2})(?:[:.](\d{1,2}))?\s*(am|pm|a\.m\.|p\.m\.)?$/i);
  if (!m) {
    // 1930 / 0730, four digits with no separator
    const military = text.match(/^(\d{2})(\d{2})$/);
    if (military) m = [text, military[1], military[2], undefined, undefined];
  }
  if (!m) return null;

  let hours = parseInt(m[1], 10);
  const minutes = parseInt(m[2], 10);
  const seconds = m[3] ? parseInt(m[3], 10) : 0;
  const meridiem = (m[4] || "").toLowerCase().replace(/\./g, "");

  if (minutes > 59 || seconds > 59) return null;
  if (meridiem === "pm" && hours < 12) hours += 12;
  if (meridiem === "am" && hours === 12) hours = 0;
  if (hours > 23) return null;

  return hours * 3600 + minutes * 60 + seconds;
}

/** Seconds since midnight -> "HH:mm:ss", matching the rest of the Log. */
function secondsToClock(totalSeconds) {
  const s = ((Math.round(totalSeconds) % 86400) + 86400) % 86400;
  const pad = (n) => (n < 10 ? "0" + n : "" + n);
  return pad(Math.floor(s / 3600)) + ":" + pad(Math.floor((s % 3600) / 60)) + ":" + pad(s % 60);
}

/**
 * Ordering key for a time of day. Anything before DAY_ROLLOVER_HOUR is
 * treated as the small hours of the following morning, so 00:30 sorts after
 * 23:50 instead of jumping to the top of an event that started at 17:00.
 */
function timeSortKey(seconds) {
  return seconds < DAY_ROLLOVER_HOUR * 3600 ? seconds + 86400 : seconds;
}

/**
 * The time right now, as seconds since midnight in venue local time.
 *
 * Built from the individual clock parts rather than a formatted string:
 * with hour12:false, some browsers write the first hour after midnight as
 * "24:15:00", which doesn't parse - so entries between midnight and 1am
 * were being stamped 00:00:00.
 */
function nowSeconds() {
  const parts = {};
  new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Perth",
    hour: "numeric",
    minute: "numeric",
    second: "numeric",
    hourCycle: "h23",
  })
    .formatToParts(new Date())
    .forEach((p) => { parts[p.type] = p.value; });
  const h = parseInt(parts.hour, 10) % 24;
  const m = parseInt(parts.minute, 10);
  const s = parseInt(parts.second, 10);
  return h * 3600 + m * 60 + s;
}

function isNotApplicable(text) {
  return /^n\/?a$/i.test(text.trim());
}

/**
 * Ribbon button handler: "Reset Milestones".
 *
 * Lets every milestone log once more on THIS PC, even though its wording is
 * already in the Log: type (or re-type) its time on Event Checks and a fresh
 * row goes in. It does not touch the Log - nothing already written is
 * changed or removed - and nothing logs until a time is actually entered.
 * Applies until each milestone has logged again, or Excel is closed.
 *
 * Also clears out what older builds left behind: the hidden "_LoggerState"
 * sheet and the saved milestone record in the add-in settings.
 */
async function resetMilestoneLog(event) {
  const addresses = Object.keys(MILESTONE_MAP);
  addresses.forEach((a) => rearmed.add(a));
  const message =
    `Milestones reset: all ${addresses.length} will log again the next time their time is entered ` +
    `on this PC (re-type a time to log it again).`;

  try {
    await Excel.run(async (context) => {
      const old = context.workbook.settings.getItemOrNullObject(OLD_MILESTONE_STATE_KEY);
      old.load("isNullObject");
      const stateSheet = context.workbook.worksheets.getItemOrNullObject(STATE_SHEET);
      stateSheet.load("isNullObject");
      await context.sync();
      if (!old.isNullObject) old.delete();
      if (!stateSheet.isNullObject) stateSheet.delete();
      await context.sync();
    });
  } catch (err) {
    // Deleting can fail on a workbook with protected structure. Harmless.
    console.error("resetMilestoneLog (tidying old state):", err);
  }

  console.log(message);
  sendToDialog({ kind: "status", ok: !/^(error|reset failed)/i.test(message), text: message });
  if (event) event.completed();
}

function openLogger(event) {
  if (dialog) {
    event.completed();
    return;
  }

  Office.context.ui.displayDialogAsync(
    DIALOG_URL,
    { height: DIALOG_HEIGHT, width: DIALOG_WIDTH, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Dialog failed to open: ${asyncResult.error.message}`);
        event.completed();
        return;
      }

      dialog = asyncResult.value;

      // Everything the dialog asks for arrives here. This page (not the
      // dialog) does all the Excel.run work.
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        await routeDialogMessage(arg.message);
      });

      // If the user closes the floating box, forget the reference so the
      // ribbon button can open a fresh one next time.
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        dialog = null;
      });

      event.completed();
    }
  );
}

/**
 * Dialog messages are JSON: { kind: "ready" | "text" | "entry", ... }.
 * A plain string (from an older cached dialog) is treated as shorthand.
 */
async function routeDialogMessage(raw) {
  let msg;
  try {
    msg = JSON.parse(raw);
  } catch (e) {
    msg = { kind: "text", text: raw };
  }

  if (msg.kind === "ready") {
    await sendMenu();
    return;
  }

  let status;
  if (msg.kind === "entry") {
    status = await withLogLock(() => handleButtonEntry(msg.fields || [], msg.startedAt));
  } else {
    status = await withLogLock(() => handleLogEntry(msg.text || "", msg.startedAt));
  }

  // No milestone catch-up here: with several people logging, every entry
  // sweeping Event Checks was another chance to log a milestone twice.
  sendToDialog({ kind: "status", ok: !/^error/i.test(status), text: status });
}

function sendToDialog(obj) {
  if (!dialog) return;
  try {
    dialog.messageChild(JSON.stringify(obj));
  } catch (err) {
    console.error("messageChild failed:", err);
  }
}

/** Reads the button rows off Inputs and sends them to the dialog in chunks. */
async function sendMenu() {
  let result = { rows: [], columns: "" };
  try {
    result = await readButtonRows();
  } catch (err) {
    console.error(err);
    sendToDialog({ kind: "status", ok: false, text: "Error reading Inputs: " + (err.message || err) });
    return;
  }

  const json = JSON.stringify(result);
  const id = Date.now().toString(36);
  const n = Math.max(1, Math.ceil(json.length / MESSAGE_CHUNK_CHARS));
  for (let i = 0; i < n; i++) {
    sendToDialog({
      kind: "menuChunk",
      id: id,
      i: i,
      n: n,
      data: json.slice(i * MESSAGE_CHUNK_CHARS, (i + 1) * MESSAGE_CHUNK_CHARS),
    });
  }
}

/** Normalises header text so "Button Row" and "button_row" match. */
function headerKey(text) {
  return (text === null || text === undefined ? "" : text.toString()).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Works out which column is which, by header text first and fixed position
 * second, and returns the indexes plus a line describing what it found (so a
 * mismatch is visible instead of silently shifting everything sideways).
 */
function mapInputsColumns(headerRow, offset) {
  const fallback = {
    code: 0,
    label: 1,
    type: INPUTS_TYPE_COL,
    group: INPUTS_GROUP_COL,
    colour: INPUTS_COLOUR_COL,
    home: INPUTS_HOME_COL,
    row: INPUTS_ROW_COL,
    parentRow: INPUTS_PARENT_ROW_COL,
  };

  const found = {};
  for (let c = 0; c < headerRow.length; c++) {
    const key = headerKey(headerRow[c]);
    if (key === "") continue;
    for (const field of Object.keys(INPUTS_HEADERS)) {
      if (found[field] === undefined && INPUTS_HEADERS[field].indexOf(key) !== -1) found[field] = c + offset;
    }
  }

  const cols = {};
  const described = [];
  for (const field of Object.keys(fallback)) {
    const byHeader = found[field] !== undefined;
    cols[field] = byHeader ? found[field] : fallback[field];
    described.push(`${field}=${columnIndexToLetter(cols[field])}${byHeader ? "" : " (assumed)"}`);
  }
  cols.description = described.join(", ");
  return cols;
}

/** Returns [code, label, type, group, colour, home, row, parentRow] for every Inputs row with a Button type. */
async function readButtonRows() {
  return await Excel.run(async (context) => {
    const inputsSheet = context.workbook.worksheets.getItemOrNullObject(INPUTS_SHEET);
    inputsSheet.load("isNullObject");
    await context.sync();
    if (inputsSheet.isNullObject) throw new Error(`"${INPUTS_SHEET}" sheet not found.`);

    const used = inputsSheet.getUsedRangeOrNullObject();
    used.load(["isNullObject", "values", "columnIndex"]);
    await context.sync();
    if (used.isNullObject) return { rows: [], columns: "" };

    // Offset in case the used range doesn't start in column A.
    const off = used.columnIndex;
    const cell = (row, col) => ((row[col - off] === undefined ? "" : row[col - off]) || "").toString().trim();

    const cols = mapInputsColumns(used.values[0] || [], off);
    console.log("Inputs columns: " + cols.description);

    const rows = [];
    for (const row of used.values) {
      const type = cell(row, cols.type);
      const label = cell(row, cols.label);
      if (!label || BUTTON_TYPES.indexOf(type.toLowerCase()) === -1) continue;
      rows.push([
        cell(row, cols.code),
        label,
        type,
        cell(row, cols.group),
        cell(row, cols.colour),
        cell(row, cols.home),
        cell(row, cols.row),
        cell(row, cols.parentRow),
      ]);
    }
    return { rows: rows, columns: cols.description };
  });
}

/** Status-line wording after a dialog entry lands in the Log. */
function entryStatus(placed, time, values) {
  if (placed.row === 0) return "Error: the Log is full.";
  const where = placed.reason && placed.reason.indexOf("inserted") === 0 ? " (slotted in by time)" : "";
  return `Row ${placed.row} at ${secondsToClock(time.seconds)}${where}: ${values.filter((v) => v !== "").join(" | ")}`;
}

/**
 * Writes a row built from dialog buttons. The fields are already full labels
 * (callsign, department, location, details), so no dictionary lookup. The
 * time is when the operator pressed the first button for this entry.
 */
async function handleButtonEntry(fields, startedAt) {
  try {
    return await Excel.run(async (context) => {
      const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
      logSheet.load("isNullObject");
      await context.sync();
      if (logSheet.isNullObject) return `Error: "${LOG_SHEET}" sheet not found.`;

      const values = [0, 1, 2, 3].map((i) => (fields[i] || "").toString().trim());
      if (values.every((v) => v === "")) return "Nothing to log.";

      const time = entryStartSeconds(startedAt);
      const placed = await writeLogRow(context, logSheet, time.seconds, values);
      if (placed.row !== 0) await refreshSecondaryValidation(context, logSheet);

      return entryStatus(placed, time, values);
    });
  } catch (err) {
    console.error(err);
    return "Error: " + (err.message || err);
  }
}

// Required for ExecuteFunction ribbon commands (enforced since Oct 2022,
// strictly so under the shared runtime) - without this, Office can't find
// openLogger by name alone and the button click does nothing.
Office.actions.associate("openLogger", openLogger);
Office.actions.associate("resetMilestoneLog", resetMilestoneLog);
// Ready for a "Sync milestones" ribbon button if you add one to the manifest.
Office.actions.associate("syncMilestones", syncMilestones);

/**
 * Parses a space-delimited shorthand line ("l1s ps ed14 ws"), resolves each
 * token against the Inputs sheet, and puts one row into the Log at the time
 * the operator started typing it.
 */
async function handleLogEntry(rawText, startedAt) {
  try {
    return await Excel.run(async (context) => {
      const inputsSheet = context.workbook.worksheets.getItemOrNullObject(INPUTS_SHEET);
      const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
      inputsSheet.load("isNullObject");
      logSheet.load("isNullObject");
      await context.sync();

      if (inputsSheet.isNullObject) return `Error: "${INPUTS_SHEET}" sheet not found.`;
      if (logSheet.isNullObject) return `Error: "${LOG_SHEET}" sheet not found.`;

      // --- Build the shorthand dictionary ---
      const inputsRange = inputsSheet.getUsedRange();
      inputsRange.load("values");
      await context.sync();

      const dictionary = {};
      for (const row of inputsRange.values || []) {
        const shorthand = (row[0] || "").toString().trim().toUpperCase();
        const expansion = (row[1] || "").toString().trim();
        if (shorthand === "CODE") continue; // column-header row
        if (shorthand !== "") dictionary[shorthand] = expansion;
      }

      // --- Split the single line on whitespace into up to 4 fields ---
      const tokens = rawText.trim().split(/\s+/).filter((t) => t !== "");
      if (tokens.length === 0) return "Nothing to log.";

      const resolved = [0, 1, 2, 3].map((i) => {
        const token = tokens[i];
        if (!token || token === "-") return "";
        const key = token.toUpperCase();
        return dictionary[key] !== undefined ? dictionary[key] : token;
      });

      // --- Slot it into the Log by the time typing started ---
      const time = entryStartSeconds(startedAt);
      const placed = await writeLogRow(context, logSheet, time.seconds, resolved);
      if (placed.row !== 0) await refreshSecondaryValidation(context, logSheet);

      return entryStatus(placed, time, resolved);
    });
  } catch (err) {
    console.error(err);
    return "Error: " + (err.message || err);
  }
}
