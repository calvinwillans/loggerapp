// Set this to match where you host these files, e.g. "https://yourname.github.io/venue-logger-dialog"
const BASE_URL = "https://calvinwillans.github.io/loggerapp";
const DIALOG_URL = BASE_URL + "/dialog.html";

const INPUTS_SHEET = "Inputs";
const LOG_SHEET = "Log";
const LOG_FIRST_DATA_ROW = 8; // first entry always lands here, matching the template
const LOG_MAX_ROW = 5000; // generous ceiling when scanning for the next empty row

let dialog = null;

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
      // Scans actual values in column A starting at LOG_FIRST_DATA_ROW, rather
      // than trusting the sheet's used range (which also counts formatting-only
      // rows, e.g. borders/fills left over from the template, and can overshoot).
      const scanRange = logSheet.getRange(`A${LOG_FIRST_DATA_ROW}:A${LOG_MAX_ROW}`);
      scanRange.load("values");
      await context.sync();

      let nextRow = LOG_MAX_ROW + 1; // fallback if every scanned row is already filled
      const colValues = scanRange.values;
      for (let i = 0; i < colValues.length; i++) {
        const cell = colValues[i][0];
        if (cell === "" || cell === null || cell === undefined) {
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
