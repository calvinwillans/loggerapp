// Set this to match where you host these files, e.g. "https://yourname.github.io/venue-logger-dialog"
const BASE_URL = "https://calvinwillans.github.io/loggerapp";
const DIALOG_URL = BASE_URL + "/dialog.html";

const INPUTS_SHEET = "Inputs";
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
 * Registers a listener on "Event Checks": when a mapped cell gets a value,
 * appends the corresponding row to Log. Each cell only fires once - once
 * logged, it's recorded on the hidden "_LoggerState" sheet so re-editing
 * that cell later (fixing a typo, say) doesn't create a duplicate row.
 * Delete the matching row on _LoggerState to let a cell fire again, or
 * clear the whole sheet to reset between events.
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

async function onMilestoneSheetChanged(eventArgs) {
  await Excel.run(async (context) => {
    const checksSheet = context.workbook.worksheets.getItem(MILESTONE_SHEET);

    const changedRange = checksSheet.getRange(eventArgs.address.split("!").pop());
    changedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
    await context.sync();

    // Which mapped cells does this edit touch?
    const touched = Object.keys(MILESTONE_MAP).filter((cellAddress) => {
      const { row, col } = a1ToRowCol(cellAddress);
      const withinRows = row >= changedRange.rowIndex && row < changedRange.rowIndex + changedRange.rowCount;
      const withinCols = col >= changedRange.columnIndex && col < changedRange.columnIndex + changedRange.columnCount;
      return withinRows && withinCols;
    });
    if (touched.length === 0) return;

    const alreadyLogged = await getLoggedMilestones(context);

    for (const cellAddress of touched) {
      if (alreadyLogged.has(cellAddress.toUpperCase())) continue;

      const cell = checksSheet.getRange(cellAddress);
      cell.load("text");
      await context.sync();
      const displayValue = (cell.text[0][0] || "").toString().trim();
      if (displayValue === "" || isNotApplicable(displayValue)) continue; // cleared or marked N/A - not populated

      await appendMilestoneRow(context, MILESTONE_MAP[cellAddress]);
      await markMilestoneLogged(context, cellAddress);
    }
  });
}

async function appendMilestoneRow(context, mapping) {
  const logSheet = context.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
  logSheet.load("isNullObject");
  await context.sync();
  if (logSheet.isNullObject) {
    console.error(`"${LOG_SHEET}" sheet not found - milestone not logged.`);
    return;
  }

  const nextRow = await findNextLogRow(context, logSheet);
  const timeString = new Date().toLocaleTimeString("en-AU", { hour12: false, timeZone: "Australia/Perth" });

  logSheet.getRange(`A${nextRow}:E${nextRow}`).values = [
    [timeString, mapping.b || "", mapping.c || "", mapping.d || "", mapping.e || ""],
  ];
  await context.sync();
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

async function getLoggedMilestones(context) {
  const stateSheet = context.workbook.worksheets.getItemOrNullObject(STATE_SHEET);
  stateSheet.load("isNullObject");
  await context.sync();
  if (stateSheet.isNullObject) return new Set();

  const used = stateSheet.getUsedRangeOrNullObject();
  used.load(["isNullObject", "values"]);
  await context.sync();
  if (used.isNullObject) return new Set();

  const logged = new Set();
  for (const row of used.values) {
    const key = (row[0] || "").toString().trim().toUpperCase();
    if (key && key !== "LOGGEDMILESTONES") logged.add(key);
  }
  return logged;
}

async function markMilestoneLogged(context, cellAddress) {
  let stateSheet = context.workbook.worksheets.getItemOrNullObject(STATE_SHEET);
  stateSheet.load("isNullObject");
  await context.sync();

  if (stateSheet.isNullObject) {
    stateSheet = context.workbook.worksheets.add(STATE_SHEET);
    stateSheet.getRange("A1").values = [["LoggedMilestones"]];
    stateSheet.visibility = Excel.SheetVisibility.hidden;
    await context.sync();
  }

  const used = stateSheet.getUsedRangeOrNullObject();
  used.load(["isNullObject", "rowIndex", "rowCount"]);
  await context.sync();

  const nextIndex = used.isNullObject ? 1 : used.rowIndex + used.rowCount; // 0-based
  stateSheet.getRangeByIndexes(nextIndex, 0, 1, 1).values = [[cellAddress.toUpperCase()]];
  await context.sync();
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
 * Ribbon button handler. Clears the hidden "_LoggerState" sheet that tracks
 * which "Event Checks" milestone cells have already fired, so any cell that
 * was cleared and repopulated (a correction) - or every cell, if the sheet
 * was just wiped for a new event - will log again next time it's edited.
 * Does NOT touch anything already written to Log.
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
      // Nothing to do if it doesn't exist yet - already reset.
    });
  } catch (err) {
    console.error("resetMilestoneLog failed:", err);
  }
  event.completed();
}

Office.actions.associate("resetMilestoneLog", resetMilestoneLog);

/**
 * Ribbon button handler. Opens the floating logger dialog. If it's already
 * open, does nothing - there's only ever one at a time.
 */
function openLogger(event) {
  if (dialog) {
    event.completed();
    return;
  }

  Office.context.ui.displayDialogAsync(
    DIALOG_URL,
    { height: 20, width: 28, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Dialog failed to open: ${asyncResult.error.message}`);
        event.completed();
        return;
      }

      dialog = asyncResult.value;

      // The dialog sends the raw shorthand line here every time the user
      // presses Enter. This page (not the dialog) does the Excel.run work.
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        const status = await handleLogEntry(arg.message);
        dialog.messageChild(status);
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

// Required for ExecuteFunction ribbon commands (enforced since Oct 2022,
// strictly so under the shared runtime) - without this, Office can't find
// openLogger by name alone and the button click does nothing.
Office.actions.associate("openLogger", openLogger);

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
