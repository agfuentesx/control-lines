const EXTRACT_WORK_SHEET_ID = "1jsc56u6raC6JjU_jLSzBqZ01DEagBnbXumqp22cojJs"

// ── Column mapping for each line (0-indexed) ──────────────────────────────────
const LINE_CONFIGS = [
  { key: "line_1", cols: { module: 0,  date: 1,  frequency: 2,  info: 3  } }, // A–D
  { key: "line_2", cols: { module: 4,  date: 5,  frequency: 6,  info: 7  } }, // E–H
  { key: "line_3", cols: { module: 8,  date: 9,  frequency: 10, info: 11 } }, // I–L
  { key: "line_4", cols: { module: 12, date: 13, frequency: 14, info: 15 } }  // M–P
]

// ── Main function ─────────────────────────────────────────────────────────────
function extractResumenToJson() {
  let spreadsheet = SpreadsheetApp.openById(EXTRACT_WORK_SHEET_ID)
  let sheet       = spreadsheet.getSheetByName("resumen")
  let lastRow     = sheet.getLastRow()

  if (lastRow < 4) {
    Logger.log("No data found from row 4 onwards.")
    return null
  }

  // Read all data from row 4 downward, columns A to P (16 columns)
  let numRows = lastRow - 3
  let data    = sheet.getRange(4, 1, numRows, 16).getValues()

  // Build the lines array following the JSON structure
  let lines = LINE_CONFIGS.map(config => {
    let entries = data
      .filter(row => row[config.cols.module] !== "" && 
                     row[config.cols.module] !== null)  // Skip empty rows per line
      .map(row => ({
        module   : String(row[config.cols.module]),
        date     : formatDate(row[config.cols.date]),
        frecuency: String(row[config.cols.frequency]),
        info     : String(row[config.cols.info])
      }))

    return { [config.key]: entries }
  })

  let result = { lines: lines }
  let json   = JSON.stringify(result, null, 2)

  Logger.log(json)
  saveJsonToDrive(json)

  return json
}

// ── Helper: format date values from the sheet ─────────────────────────────────
function formatDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy")
  }
  return value !== null && value !== undefined ? String(value) : ""
}

// ── Helper: save the JSON as a file in Google Drive ───────────────────────────
function saveJsonToDrive(jsonString) {
  let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss")
  let fileName  = "resumen_" + timestamp + ".json"
  let file      = DriveApp.createFile(fileName, jsonString, MimeType.PLAIN_TEXT)

  Logger.log("✅ JSON saved to Drive: " + file.getUrl())
}
