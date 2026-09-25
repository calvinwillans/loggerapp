// Set this to match where you host these files, e.g. "https://yourname.github.io/venue-logger-dialog"
const BASE_URL = "https://calvinwillans.github.io/loggerapp";
const DIALOG_URL = BASE_URL + "/dialog.html";

const INPUTS_SHEET = "Inputs";
// Inputs columns used to build the dialog's buttons (A = code, B = label).
const INPUTS_TYPE_COL = 2;   // C - Callsign / Department / Incident / Location
const INPUTS_GROUP_COL = 3;  // D - group path, e.g. "Blocks/200s" or "Level 2;Toilets"
const INPUTS_COLOUR_COL = 4; // E - hex colour, e.g. #1D4ED8
const INPUTS_HOME_COL = 5;   // F - callsigns only: a location label, or a location folder path
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

const MILESTONE_SHEET = "Event Checks";
const STATE_SHEET = "_LoggerState";

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
  C104: { b: "Base", c: "All Call", d: "", e: "Level 4 whitelevel checks completed" },
  C105: { b: "Base", c: "All Call", d: "", e: "Level 3 whitelevel checks completed" },
  C106: { b: "Base", c: "All Call", d: "", e: "Level 2 whitelevel checks completed" },
  C107: { b: "Base", c: "All Call", d: "", e: "Level 1 whitelevel checks completed" },
  C108: { b: "Base", c: "All Call", d: "", e: "Ground floor whitelevel checks completed" },
  C109: { b: "Base", c: "All Call", d: "", e: "Venue clear. Switch to non-event channels" },
  C115: { b: "Base", c: "", d: "", e: "External checks complete. Channel 1 handover to CRO & EWIS back to auto" },
};

let dialog = null;

// Runs once, as soon as this page loads. With the shared runtime + a "long"
// lifetime in the manifest, that's now on workbook open rather than only
// when the ribbon button is clicked or a dialog message arrives.
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    try {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    } catch (err) {
      console.error("setStartupBehavior failed:", err);
    }
    await registerDepartmentWatcher();
    await registerMilestoneWatcher();
    const opening = await sweepMilestones();
    console.log("On open: " + sweepSummary(opening));
  }
});

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
  await Excel.run(async (context) => {
    const logSheet = context.workbook.worksheets.getItem(LOG_SHEET);

    const changedRange = logSheet.getRange(eventArgs.address.split("!").pop());
    changedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
    await context.sync();

    const deptColIndex = columnLetterToIndex(DEPARTMENT_COLUMN);
    const touchesDept =
      changedRange.columnIndex <= deptColIndex &&
      deptColIndex < changedRange.columnIndex + changedRange.columnCount;
    if (!touchesDept) return;

    const firstDataRowIndex = LOG_FIRST_DATA_ROW - 1;
    const startRow = Math.max(changedRange.rowIndex, firstDataRowIndex);
    const endRow = changedRange.rowIndex + changedRange.rowCount - 1;
    if (endRow < firstDataRowIndex) return;

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
async function sweepMilestones() {
  try {
    return await Excel.run(async (context) => {
      const checksSheet = context.workbook.worksheets.getItemOrNullObject(MILESTONE_SHEET);
      const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
      checksSheet.load("isNullObject");
      logSheet.load("isNullObject");
      await context.sync();

      if (checksSheet.isNullObject) return { checked: 0, logged: 0, error: `"${MILESTONE_SHEET}" sheet not found.` };
      if (logSheet.isNullObject) return { checked: 0, logged: 0, error: `"${LOG_SHEET}" sheet not found.` };

      // Read every mapped cell in one round trip.
      const addresses = Object.keys(MILESTONE_MAP);
      const cells = addresses.map((a) => {
        const r = checksSheet.getRange(a);
        r.load("text");
        return r;
      });

      // Details already logged, so nothing gets written twice.
      const detailsRange = logSheet.getRange(`E${LOG_FIRST_DATA_ROW}:E${LOG_MAX_ROW}`);
      detailsRange.load("text");
      const timeRange = logSheet.getRange(`A${LOG_FIRST_DATA_ROW}:A${LOG_MAX_ROW}`);
      timeRange.load("text");
      await context.sync();

      const already = new Set();
      for (const row of detailsRange.text) {
        const text = (row[0] || "").toString().trim().toLowerCase();
        if (text !== "") already.add(text);
      }

      let nextRow = LOG_MAX_ROW + 1;
      const timeText = timeRange.text;
      for (let i = 0; i < timeText.length; i++) {
        if ((timeText[i][0] || "").toString().trim() === "") { nextRow = LOG_FIRST_DATA_ROW + i; break; }
      }

      const timeString = new Date().toLocaleTimeString("en-AU", { hour12: false, timeZone: "Australia/Perth" });
      let logged = 0;

      addresses.forEach((address, i) => {
        const displayValue = (cells[i].text[0][0] || "").toString().trim();
        if (displayValue === "" || isNotApplicable(displayValue)) return; // not filled in yet

        const mapping = MILESTONE_MAP[address];
        const details = (mapping.e || "").trim();
        if (details === "" || already.has(details.toLowerCase())) return; // already in the Log

        if (nextRow > LOG_MAX_ROW) return; // Log is full

        logSheet.getRange(`A${nextRow}:E${nextRow}`).values = [
          [timeString, mapping.b || "", mapping.c || "", mapping.d || "", details],
        ];
        already.add(details.toLowerCase());
        nextRow++;
        logged++;
      });

      if (logged > 0) await context.sync();
      return { checked: addresses.length, logged: logged, error: null };
    });
  } catch (err) {
    console.error("sweepMilestones failed:", err);
    return { checked: 0, logged: 0, error: err.message || String(err) };
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

function isNotApplicable(text) {
  return /^n\/?a$/i.test(text.trim());
}

/**
 * Ribbon button handler, kept for the existing button. Milestones are now
 * deduped against the Log itself, so there is nothing to reset: to make a
 * milestone log again, clear its Details text from the Log. This also
 * removes the old hidden "_LoggerState" sheet if one is still lying around.
 */
async function resetMilestoneLog(event) {
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
    // Deleting can fail if the workbook structure is protected. Harmless now.
    console.error("resetMilestoneLog:", err);
  }
  const result = await sweepMilestones();
  console.log("Reset: " + sweepSummary(result));
  sendToDialog({ kind: "status", ok: !result.error, text: sweepSummary(result) });
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
  let rows = [];
  try {
    rows = await readButtonRows();
  } catch (err) {
    console.error(err);
    sendToDialog({ kind: "status", ok: false, text: "Error reading Inputs: " + (err.message || err) });
    return;
  }

  const json = JSON.stringify(rows);
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

/** Returns [code, label, type, group, colour, home] for every Inputs row with a Button type. */
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

    const rows = [];
    for (const row of used.values) {
      const type = cell(row, INPUTS_TYPE_COL);
      const label = cell(row, 1);
      if (!label || BUTTON_TYPES.indexOf(type.toLowerCase()) === -1) continue;
      rows.push([
        cell(row, 0),
        label,
        type,
        cell(row, INPUTS_GROUP_COL),
        cell(row, INPUTS_COLOUR_COL),
        cell(row, INPUTS_HOME_COL),
      ]);
    }
    return rows;
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

      const summary = resolved.filter((v) => v !== "").join(" | ");
      return `Row ${nextRow}: ${summary}`;
    });
  } catch (err) {
    console.error(err);
    return "Error: " + (err.message || err);
  }
}
