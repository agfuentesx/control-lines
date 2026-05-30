// ── Constants ─────────────────────────────────────────────────────────────────
const FORM_VIEW_URL  = "https://docs.google.com/forms/d/e/1FAIpQLSfDDkoEfZD9lg2udMloIbloUULYXlafTDxIrSIGoCt8e3aKGA/viewform?embedded=true"
const ADMIN_EMAILS   = ["angelfuentesemiliani@gmail.com"]
//const EXTRACT_WORK_SHEET_ID = "1jsc56u6raC6JjU_jLSzBqZ01DEagBnbXumqp22cojJs"

// ── Entry point ───────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Control Lines')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

// ── Called by the frontend to get the active user's role ──────────────────────
function checkUserRole(email) {
  const clean = email.toLowerCase().trim()
  const role  = ADMIN_EMAILS.map(e => e.toLowerCase()).includes(clean) 
                ? 'admin' 
                : 'visitante'
  Logger.log("checkUserRole → " + clean + " | " + role)
  return { email: clean, role }
}

// ── Called by the frontend to get module data as JSON ─────────────────────────
// Reads the spreadsheet directly — does NOT save to Drive (use extractResumenToJson for that)
function getResumenData() {
  const spreadsheet = SpreadsheetApp.openById(EXTRACT_WORK_SHEET_ID)
  const sheet       = spreadsheet.getSheetByName("resumen")
  const lastRow     = sheet.getLastRow()

  if (lastRow < 4) return JSON.stringify({ lines: [] })

  const data = sheet.getRange(4, 1, lastRow - 3, 16).getValues()

  const LINE_CONFIGS = [
    { key: "line_1", cols: { module: 0,  date: 1,  frequency: 2,  info: 3  } },
    { key: "line_2", cols: { module: 4,  date: 5,  frequency: 6,  info: 7  } },
    { key: "line_3", cols: { module: 8,  date: 9,  frequency: 10, info: 11 } },
    { key: "line_4", cols: { module: 12, date: 13, frequency: 14, info: 15 } }
  ]

  const lines = LINE_CONFIGS.map(config => {
    const entries = data
      .filter(row => row[config.cols.module] !== "" && row[config.cols.module] !== null)
      .map(row => ({
        module   : String(row[config.cols.module]),
        date     : formatDate(row[config.cols.date]),   // formatDate lives in extractInformation.gs
        frecuency: String(row[config.cols.frequency]),
        info     : String(row[config.cols.info])
      }))
    return { [config.key]: entries }
  })

  return JSON.stringify({ lines })
}

// ── Called by the frontend to get the form embed URL ─────────────────────────
function getFormUrl() {
  return FORM_VIEW_URL
}
