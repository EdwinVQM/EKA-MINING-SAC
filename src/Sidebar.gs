const SIDEBAR_TITLE = 'Panel de pruebas';
const FALLBACK_SIDEBAR_HTML = `<!DOCTYPE html>
<html lang="es">
  <head>
    <meta charset="utf-8" />
    <style>
      body {
        font-family: "Segoe UI", Roboto, sans-serif;
        margin: 0;
        padding: 16px;
        background: #f5f7fa;
        color: #111827;
      }
      h1 {
        font-size: 18px;
        margin: 0 0 12px;
      }
      .box {
        background: #ffffff;
        border-radius: 8px;
        padding: 12px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.12);
      }
      .error {
        color: #b91c1c;
        font-weight: 600;
      }
      ol {
        padding-left: 18px;
      }
      code {
        background: #e0e7ff;
        padding: 2px 4px;
        border-radius: 4px;
      }
    </style>
    <title>Panel de pruebas</title>
  </head>
  <body>
    <div class="box">
      <h1>Panel no disponible</h1>
      <p class="error">No se pudo cargar <code>Sidebar.html</code>.</p>
      <p>Verifica que el archivo exista en el proyecto de Apps Script y vuelve a ejecutar el menú.</p>
      <p><strong>Error devuelto:</strong> {{error}}</p>
      <p>Hoja activa: <strong>{{sheetName}}</strong></p>
      <ol>
        <li>Confirma que el repositorio contenga <code>src/Sidebar.html</code>.</li>
        <li>Ejecuta <code>clasp push</code> desde la carpeta del proyecto (donde está <code>.clasp.json</code>).</li>
        <li>Recarga la hoja de cálculo y vuelve a abrir el sidebar.</li>
      </ol>
    </div>
  </body>
</html>`;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧪 Pruebas Sidebar')
    .addItem('Abrir panel', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const ui = SpreadsheetApp.getUi();
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  let htmlOutput;

  try {
    const template = HtmlService.createTemplateFromFile('Sidebar');
    template.sheetName = sheetName;
    htmlOutput = template.evaluate();
  } catch (error) {
    console.error('No se pudo cargar Sidebar.html', error);
    const message = error && error.message ? error.message : String(error);
    const html = FALLBACK_SIDEBAR_HTML
      .replace('{{error}}', message)
      .replace('{{sheetName}}', sheetName);
    htmlOutput = HtmlService.createHtmlOutput(html);
  }

  htmlOutput.setTitle(SIDEBAR_TITLE).setWidth(320);
  ui.showSidebar(htmlOutput);
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
