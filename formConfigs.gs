const FORM_URL = "https://docs.google.com/forms/d/1lT0MG3nAxkMpqu5zhpfKlXmNF-fRzmSGYIiIqxlRojM/edit"
const WORK_SHEET_ID = "1eB3nGVkZSkJiUnYv5L8s5SmFZgxlhaVXuUgRV65w-ak"

let form = FormApp.openByUrl(FORM_URL)

function setupTrigger() {
  // Ejecuta esta función UNA SOLA VEZ para instalar el trigger
  let form = FormApp.openByUrl(FORM_URL);
  ScriptApp.newTrigger('saveResponseLongFormat')
    .forForm(form)
    .onFormSubmit()
    .create();
}

function saveResponseLongFormat(e) {
  let response = e.response;
  let items = response.getItemResponses();
  
  // Abre la hoja donde guardaremos en formato largo
  let ss = SpreadsheetApp.openById(WORK_SHEET_ID);
  let sheet = ss.getSheetByName("respuestas_formato_largo") || 
              ss.insertSheet("respuestas_formato_largo");
  
  // Agrega encabezados si la hoja está vacía
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([[
      "ID de respuesta", "Timestamp", "Pregunta", "Respuesta"
    ]]);
  }
  
  let timestamp = response.getTimestamp();
  let responseId = response.getId();
  let output = [];
  
  // Una fila por cada pregunta/respuesta
  items.forEach(item => {
    output.push([
      responseId,
      timestamp,
      item.getItem().getTitle(),  // Texto de la pregunta
      item.getResponse()           // Respuesta
    ]);
  });
  
  if (output.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, output.length, 4).setValues(output);
  }
}

function reset_form() {
  let items = form.getItems()
  items.forEach(item => form.deleteItem(item))
}

function conditional_form() {
  reset_form()

  let work_sheet = SpreadsheetApp.openById(WORK_SHEET_ID)
  let spread_sheet = work_sheet.getSheetByName("lineas_modulos")
  let table = spread_sheet.getDataRange().getValues()
  table.shift() // Remove header row

  // ── Unique Lines from column A ──────────────────────────────
  let single_line_options = []
  table.map(row => row[0]).forEach(option => {
    if (single_line_options.indexOf(option) == -1)
      single_line_options.push(option)
  })

  // ── Unique Frequency options from column C ──────────────────
  let frequency_options = []
  table.forEach(row => {
    let freq = String(row[2]) // Column C → Frecuencia
    if (freq && freq !== "undefined" && frequency_options.indexOf(freq) == -1)
      frequency_options.push(freq)
  })

  // ── Main Line question (first page) ────────────────────────
  let line_question = form.addMultipleChoiceItem().setTitle("Linea")
  let final_line_options = []

  single_line_options.forEach(line_option => {

    // Section (page break) for this line
    let section = form.addPageBreakItem().setTitle(line_option)
    section.setGoToPage(FormApp.PageNavigationType.SUBMIT)

    // 1. Module question — choices filtered by current line
    let module_options = table
      .filter(row => row[0] == line_option)
      .map(row => row[1])
    form.addMultipleChoiceItem()
      .setTitle("Modulo")
      .setChoiceValues(module_options)

    // 2. Maintenance Date — date picker field
    form.addDateItem()
      .setTitle("Fecha de mantenimiento")

    // 3. Frequency — choices pulled from column C
    form.addMultipleChoiceItem()
      .setTitle("Frecuencia")
      .setChoiceValues(frequency_options)

    // 4. Additional Information — free text field
    form.addParagraphTextItem()
      .setTitle("Informacion adicional")

    // Link this line option to its section
    final_line_options.push(line_question.createChoice(line_option, section))
  })

  line_question.setChoices(final_line_options)
}
