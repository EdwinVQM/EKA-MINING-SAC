const SIDEBAR_TITLE = 'Panel de pruebas';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧪 Pruebas Sidebar')
    .addItem('Abrir panel', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.sheetName = SpreadsheetApp.getActiveSheet().getName();
  const htmlOutput = template
    .evaluate()
    .setTitle(SIDEBAR_TITLE)
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getActiveRowData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const row = activeCell.getRow();
  const values = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  return { row, values };
}

function logNoteOnRow(note) {
  if (typeof note !== 'string') {
    throw new Error('El texto de la nota debe ser una cadena.');
  }
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  cell.setNote(note);
  return {
    row: cell.getRow(),
    column: cell.getColumn(),
    note,
  };
}
