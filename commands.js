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

const MILESTONE_SHEET = "Event Checks";
const STATE_SHEET = "_LoggerState";

// Which milestones have been logged, and what their cells held when the
// Reset button was last pressed. Kept in the workbook's add-in settings
// rather than on a sheet, so there's nothing to hide, nothing to protect
// against, and nothing for anyone to delete by accident.
const MILESTONE_STATE_KEY = "VenueLogger.milestoneState";

// Cell on "Event Checks" -> the row to append to Log once that cell is
// populated. Add more entries here in the same shape as they come up -
// leave any column "" to leave it blank on the logged row.
const MILESTONE_MAP = {
  F23: { b: "Base", c: "All Call", d: "", e: "All departments switch over to event channels" },
  F25: { b: "Base", c: "", d: "", e: "Fire panel in event mode. EWIS switched to manual" },
  E50: { b: "Base", c: "", d: "", e: "WAPOL arrived onsite" },
  E52: { b: "Base", c: "All Call", d: "", e: "RAC Local Lounge now open" },
  E63: { b: "Base", c: "All Call", d: "", e: "External doors now open" },
  E65: { b: "Base", c: "All Call", d: "", e: "Internal doors now open" },
  C73: { b: "Base", c: "All Call", d: "", e: "Start of support act" },
  D73: { b: "Base", c: "All Call", d: "", e: "End of support act" },
  C74: { b: "Base", c: "All Call", d: "", e: "Start of main act/game" },
  C75: { b: "Base", c: "All Call", d: "", e: "Start of Intermission/Halftime" },
  D75: { b: "Base", c: "All Call", d: "", e: "End of Intermission/Halftime" },
  C87: { b: "Base", c: "All Call", d: "", e: "End of show/game. Prepare for egress" },
  C104: { b: "Security", c: "Base", d: "", e: "Level 4 whitelevel checks completed" },
  C105: { b: "Security", c: "Base", d: "", e: "Level 3 whitelevel checks completed" },
  C106: { b: "Security", c: "Base", d: "", e: "Level 2 whitelevel checks completed" },
  C107: { b: "Security", c: "Base", d: "", e: "Level 1 whitelevel checks completed" },
  C108: { b: "Security", c: "Base", d: "", e: "Ground floor whitelevel checks completed" },
  C109: { b: "Base", c: "All Call", d: "", e: "Venue clear. Switch to non-event channels" },
  C115: { b: "Security", c: "Base", d: "", e: "External checks complete. Channel 1 handover to CRO & EWIS back to auto" },
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
    const opening = await sweepMilestones();
    console.log("On open: " + sweepSummary(opening));
    await refreshSecondaryOnOpen();
  }
});

/**
 * Registers a listener on the Log sheet: any edit to column C (Department)
 * refreshes that row's Action Taken (F) dropdown to match the corresponding
 * column on the Actions sheet. Fully open-ended - add a new column to
 * Actions with any header text and it's usable immediately, no other
 * changes needed.
 */
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

async function onMilestoneSheetChanged() {
  const result = await sweepMilestones();
  if (result.error) console.error("Milestone sweep failed:", result.error);
}

/**
 * One comparable form for a Details string: non-breaking spaces become
 * ordinary ones, runs of whitespace collapse, case is ignored. Without this,
 * a double space or a stray nbsp makes a row look "not yet logged".
 */
function normaliseDetails(value) {
  return (value === null || value === undefined ? "" : value.toString())
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// Excel fires onChanged for every edit, and a recalculation on "Event Checks"
// can fire it again while a sweep is still running. Two sweeps in flight at
// once both read the Log before either writes, so both decide the same
// milestone is missing and both write it. These guards make the sweep
// single-file: concurrent callers share the running sweep, and one extra pass
// runs afterwards so a change made mid-sweep still gets picked up.
let sweepInFlight = null;
let sweepRequestedAgain = false;

async function sweepMilestones() {
  if (sweepInFlight) {
    sweepRequestedAgain = true;
    return sweepInFlight;
  }

  sweepInFlight = runMilestoneSweep();
  let result;
  try {
    result = await sweepInFlight;
  } finally {
    sweepInFlight = null;
  }

  if (sweepRequestedAgain) {
    sweepRequestedAgain = false;
    const second = await sweepMilestones();
    return {
      checked: second.checked,
      logged: result.logged + second.logged,
      error: result.error || second.error,
      report: second.report,
    };
  }
  return result;
}

/**
 * Checks EVERY cell in MILESTONE_MAP and logs any that are filled in but not
 * yet in the Log. Deliberately independent of the change event: it also runs
 * on workbook open, from the ribbon, and after each dialog entry, so a
 * milestone filled by formula, pasted in, or entered while the add-in was
 * asleep still gets picked up.
 *
 * Duplicates are prevented by matching against the Details text already in
 * the Log (column E), so there is no hidden state sheet to create, lose, or
 * reset - which also means it works in a workbook with protected structure.
 */
/**
 * Reads the milestone state: { logged, baseline }.
 *   logged   - addresses already written to the Log since the last reset
 *   baseline - what each cell held at the moment of that reset, so a cell
 *              that is simply still sitting there doesn't log all over again
 * "available: false" means the settings API wouldn't answer, in which case
 * the sweep falls back to matching against the Log's own contents.
 */
async function loadMilestoneState(context) {
  try {
    const setting = context.workbook.settings.getItemOrNullObject(MILESTONE_STATE_KEY);
    setting.load("isNullObject, value");
    await context.sync();

    if (setting.isNullObject) return { logged: {}, baseline: {}, available: true };

    const parsed = JSON.parse(setting.value || "{}");
    return {
      logged: parsed.logged || {},
      baseline: parsed.baseline || {},
      available: true,
    };
  } catch (err) {
    console.error("Milestone state unavailable, falling back to the Log:", err);
    return { logged: {}, baseline: {}, available: false };
  }
}

async function saveMilestoneState(context, state) {
  if (!state.available) return;
  try {
    context.workbook.settings.add(
      MILESTONE_STATE_KEY,
      JSON.stringify({ logged: state.logged, baseline: state.baseline })
    );
    await context.sync();
  } catch (err) {
    console.error("Could not save milestone state:", err);
  }
}

/** What a milestone cell currently holds, as one comparable string. */
function cellSignature(shown, held) {
  const a = shown === null || shown === undefined ? "" : shown.toString().trim();
  const b = held === null || held === undefined ? "" : held.toString().trim();
  return a !== "" ? a : b;
}

/**
 * Works out where a milestone row belongs in the Log, given its time.
 *
 * Returns { row, insert }. "insert" means an existing row has a later time,
 * so a blank row has to be pushed in at that position; otherwise the row is
 * simply the first empty one at the bottom.
 */
async function planMilestonePlacement(context, logSheet, seconds) {
  const range = logSheet.getRange(`A${LOG_FIRST_DATA_ROW}:A${LOG_MAX_ROW}`);
  range.load("values, text");
  await context.sync();

  const values = range.values;
  const texts = range.text;
  const key = timeSortKey(seconds);

  for (let i = 0; i < values.length; i++) {
    const rowNumber = LOG_FIRST_DATA_ROW + i;
    const held = values[i][0];
    const shown = texts[i][0];
    const blank =
      (held === null || held === undefined || held.toString().trim() === "") &&
      (shown === null || shown === undefined || shown.toString().trim() === "");

    if (blank) return { row: rowNumber, insert: false }; // reached the bottom

    // Prefer the stored value, fall back to what's displayed.
    let existing = parseTimeToSeconds(held);
    if (existing === null) existing = parseTimeToSeconds(shown);
    if (existing === null) continue; // not a time - leave that row alone

    if (timeSortKey(existing) > key) return { row: rowNumber, insert: true };
  }

  return { row: LOG_MAX_ROW + 1, insert: false };
}

/**
 * Puts one milestone row into the Log at the right place. Returns the row it
 * landed on, or 0 if the Log is full.
 */
async function writeMilestoneRow(context, logSheet, mapping, seconds) {
  const plan = await planMilestonePlacement(context, logSheet, seconds);
  if (plan.row > LOG_MAX_ROW) return 0;

  let row = plan.row;
  if (plan.insert) {
    try {
      logSheet.getRange(`${row}:${row}`).insert(Excel.InsertShiftDirection.down);
      await context.sync();
    } catch (err) {
      // Inserting can be refused (a protected sheet, say). Rather than lose
      // the entry, fall back to adding it at the bottom, out of order.
      console.error("Could not insert a row for a backfilled milestone:", err);
      const fallback = await planMilestonePlacement(context, logSheet, 24 * 3600 - 1);
      row = fallback.row;
      if (row > LOG_MAX_ROW) return 0;
    }
  }

  logSheet.getRange(`A${row}:E${row}`).values = [
    [secondsToClock(seconds), mapping.b || "", mapping.c || "", mapping.d || "", (mapping.e || "").trim()],
  ];
  await context.sync();
  return row;
}

async function runMilestoneSweep() {
  try {
    return await Excel.run(async (context) => {
      const checksSheet = context.workbook.worksheets.getItemOrNullObject(MILESTONE_SHEET);
      const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
      checksSheet.load("isNullObject");
      logSheet.load("isNullObject");
      await context.sync();

      if (checksSheet.isNullObject) return { checked: 0, logged: 0, error: `"${MILESTONE_SHEET}" sheet not found.` };
      if (logSheet.isNullObject) return { checked: 0, logged: 0, error: `"${LOG_SHEET}" sheet not found.` };

      const state = await loadMilestoneState(context);

      // Read every mapped cell in one round trip.
      const addresses = Object.keys(MILESTONE_MAP);
      const cells = addresses.map((a) => {
        const r = checksSheet.getRange(a);
        r.load("text, values");  // text for a real time, values as the backstop
        return r;
      });

      // What's already in the Log, so nothing gets written twice. This reads
      // VALUES, not text: text is what the cell displays, and Excel clips that
      // to the column width, so a long Details string came back cut short and
      // never matched the full string being compared against it - which is why
      // the longer milestones logged again on every sweep.
      const detailsRange = logSheet.getRange(`E${LOG_FIRST_DATA_ROW}:E${LOG_MAX_ROW}`);
      detailsRange.load("values");
      await context.sync();

      const already = new Set();
      for (const row of detailsRange.values) {
        const text = normaliseDetails(row[0]);
        if (text !== "") already.add(text);
      }

      let logged = 0;
      const report = [];

      // First pass decides what needs logging, reading only.
      const toLog = [];
      addresses.forEach((address, i) => {
        // A cell counts as filled if either what it shows or what it holds is
        // non-blank, so a time entered any which way is recognised.
        const shown = (cells[i].text[0][0] === null || cells[i].text[0][0] === undefined ? "" : cells[i].text[0][0]).toString().trim();
        const held = (cells[i].values[0][0] === null || cells[i].values[0][0] === undefined ? "" : cells[i].values[0][0]).toString().trim();
        const seen = shown !== "" ? shown : held;

        if (seen === "" || isNotApplicable(shown) || isNotApplicable(held)) {
          report.push(`${address}: empty, skipped`);
          return;
        }

        const mapping = MILESTONE_MAP[address];
        const details = (mapping.e || "").trim();
        const key = normaliseDetails(details);
        if (key === "") return;

        // Already logged since the last reset.
        if (state.logged[address]) {
          report.push(`${address}: "${seen}" already logged, skipped`);
          return;
        }

        // Reset takes a snapshot of every cell. A cell still holding exactly
        // what it held then hasn't been repopulated, so it stays put; change
        // it, or clear and re-enter it, and it logs again.
        const signature = cellSignature(shown, held);
        const baseline = state.baseline[address];
        const unchangedSinceReset = baseline !== undefined && baseline === signature;
        if (unchangedSinceReset) {
          report.push(`${address}: "${seen}" unchanged since reset, skipped`);
          return;
        }

        // No reset has touched this one, so fall back to the Log itself: if
        // its wording is already there, it was logged before.
        if (baseline === undefined && already.has(key)) {
          state.logged[address] = true;
          report.push(`${address}: "${seen}" already in the Log, skipped`);
          return;
        }

        // The time comes from the cell itself, not the clock, so a milestone
        // backfilled an hour later still carries the time it happened. The
        // stored value is tried first, then what the cell displays.
        let seconds = parseTimeToSeconds(cells[i].values[0][0]);
        if (seconds === null) seconds = parseTimeToSeconds(shown);
        const parsed = seconds !== null;
        if (!parsed) seconds = nowSeconds();

        toLog.push({ address: address, mapping: mapping, details: details, seconds: seconds, parsed: parsed, seen: seen });
        already.add(key);
        state.logged[address] = true;
        delete state.baseline[address]; // it has logged since the reset
      });

      // Second pass writes them, each slotted in by its own time.
      for (const entry of toLog) {
        const row = await writeMilestoneRow(context, logSheet, entry.mapping, entry.seconds);
        if (row === 0) {
          report.push(`${entry.address}: Log is full`);
          continue;
        }
        report.push(
          `${entry.address}: "${entry.seen}" -> row ${row} at ${secondsToClock(entry.seconds)}` +
            (entry.parsed ? "" : " (no time found in the cell, used the clock)")
        );
        logged++;
      }

      if (logged > 0) {
        await saveMilestoneState(context, state);
        await context.sync();
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

/** Ribbon button: check the milestones now and report in the console. */
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

function a1ToRowCol(address) {
  const match = address.toUpperCase().match(/^([A-Z]+)(\d+)$/);
  const row = parseInt(match[2], 10) - 1;
  let col = 0;
  for (const ch of match[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row, col: col - 1 };
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

/** The time right now, as seconds since midnight in venue local time. */
function nowSeconds() {
  const parts = new Date().toLocaleTimeString("en-AU", {
    hour12: false,
    timeZone: "Australia/Perth",
  });
  const parsed = parseTimeToSeconds(parts);
  return parsed === null ? 0 : parsed;
}

function isNotApplicable(text) {
  return /^n\/?a$/i.test(text.trim());
}

/**
 * Ribbon button handler: "Reset Milestones".
 *
 * Clears the record of which milestones have been logged, so every one of
 * them can log again. It does NOT touch the Log - nothing already written is
 * changed or removed - and it does not re-log what is sitting on "Event
 * Checks" right now: a snapshot is taken of every mapped cell, and a
 * milestone logs again once its cell is repopulated (cleared and re-entered,
 * or changed to a different time).
 *
 * Also removes the old hidden "_LoggerState" sheet if one is still around
 * from an earlier version.
 */
async function resetMilestoneLog(event) {
  let message;
  try {
    message = await Excel.run(async (context) => {
      const checksSheet = context.workbook.worksheets.getItemOrNullObject(MILESTONE_SHEET);
      checksSheet.load("isNullObject");
      await context.sync();
      if (checksSheet.isNullObject) return `Error: "${MILESTONE_SHEET}" sheet not found.`;

      const addresses = Object.keys(MILESTONE_MAP);
      const cells = addresses.map((a) => {
        const r = checksSheet.getRange(a);
        r.load("text, values");
        return r;
      });
      await context.sync();

      // Snapshot what each cell holds right now; nothing re-logs until one of
      // them changes.
      const baseline = {};
      let filled = 0;
      addresses.forEach((address, i) => {
        const signature = cellSignature(cells[i].text[0][0], cells[i].values[0][0]);
        baseline[address] = signature;
        if (signature !== "") filled++;
      });

      const state = { logged: {}, baseline: baseline, available: true };
      await saveMilestoneState(context, state);

      return (
        `Milestones reset: all ${addresses.length} can log again. ` +
        `${filled} cell${filled === 1 ? "" : "s"} already filled in will log when changed or re-entered.`
      );
    });
  } catch (err) {
    console.error("resetMilestoneLog failed:", err);
    message = "Reset failed: " + (err.message || err);
  }

  // Tidy up the hidden sheet an older build used for this.
  try {
    await Excel.run(async (context) => {
      const stateSheet = context.workbook.worksheets.getItemOrNullObject(STATE_SHEET);
      stateSheet.load("isNullObject");
      await context.sync();
      if (!stateSheet.isNullObject) {
        stateSheet.delete();
        await context.sync();
      }
    });
  } catch (err) {
    // Deleting can fail on a workbook with protected structure. Harmless.
    console.error("resetMilestoneLog (old state sheet):", err);
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
    status = await handleButtonEntry(msg.fields || []);
  } else {
    status = await handleLogEntry(msg.text || "");
  }

  // Catch up on any milestone filled in since the last entry.
  const sweep = await sweepMilestones();
  if (sweep.logged > 0 || sweep.error) status += "  |  " + sweepSummary(sweep);

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
    if (used.isNullObject) return [];

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

/**
 * Writes a row built from dialog buttons. The fields are already full labels
 * (callsign, department, location, details), so no dictionary lookup.
 */
async function handleButtonEntry(fields) {
  try {
    return await Excel.run(async (context) => {
      const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
      logSheet.load("isNullObject");
      await context.sync();
      if (logSheet.isNullObject) return `Error: "${LOG_SHEET}" sheet not found.`;

      const values = [0, 1, 2, 3].map((i) => (fields[i] || "").toString().trim());
      if (values.every((v) => v === "")) return "Nothing to log.";

      const nextRow = await findNextLogRow(context, logSheet);
      const timeString = new Date().toLocaleTimeString("en-AU", {
        hour12: false,
        timeZone: "Australia/Perth",
      });

      logSheet.getRange(`A${nextRow}:E${nextRow}`).values = [[timeString, values[0], values[1], values[2], values[3]]];
      await context.sync();
      await refreshSecondaryValidation(context, logSheet);

      return `Row ${nextRow}: ${values.filter((v) => v !== "").join(" | ")}`;
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
 * token against the Inputs sheet, and appends one row to Log.
 */
async function handleLogEntry(rawText) {
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

      // --- Find the next empty row on Log ---
      // Scans the DISPLAYED text in column A starting at LOG_FIRST_DATA_ROW,
      // rather than raw values or the sheet's used range. A cell can look
      // blank on screen while its underlying value isn't a true empty string
      // (a stray space, or a template formula/format that hides a result) -
      // .text reflects what's actually visible, matching what the row looks
      // like to a person reading the sheet.
      const scanRange = logSheet.getRange(`A${LOG_FIRST_DATA_ROW}:A${LOG_MAX_ROW}`);
      scanRange.load("text");
      await context.sync();

      let nextRow = LOG_MAX_ROW + 1; // fallback if every scanned row is already filled
      const colText = scanRange.text;
      for (let i = 0; i < colText.length; i++) {
        const cellText = (colText[i][0] || "").toString().trim();
        if (cellText === "") {
          nextRow = LOG_FIRST_DATA_ROW + i;
          break;
        }
      }

      // --- Write timestamp + the 4 resolved values ---
      const timeString = new Date().toLocaleTimeString("en-AU", {
        hour12: false,
        timeZone: "Australia/Perth",
      });

      logSheet
        .getRange(`A${nextRow}:E${nextRow}`)
        .values = [[timeString, resolved[0], resolved[1], resolved[2], resolved[3]]];
      await context.sync();
      await refreshSecondaryValidation(context, logSheet);

      const summary = resolved.filter((v) => v !== "").join(" | ");
      return `Row ${nextRow}: ${summary}`;
    });
  } catch (err) {
    console.error(err);
    return "Error: " + (err.message || err);
  }
}
