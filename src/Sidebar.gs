const SIDEBAR_TITLE = 'Panel de pruebas';
const SIDEBAR_FILE = 'Sidebar';
/**
 * Plantilla HTML utilizada cuando no se encuentra la vista principal.
 * Se mantiene inline para asegurar que siempre exista un panel mínimo.
 */
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
      code,
      .template {
        background: #e0e7ff;
        padding: 2px 4px;
        border-radius: 4px;
      }
      .template {
        font-family: 'Fira Code', 'Roboto Mono', monospace;
      }
    </style>
    <title>Panel de pruebas</title>
  </head>
  <body>
    <div class="box">
      <h1>Panel no disponible</h1>
      <p class="error">No se pudo cargar la plantilla <span class="template">{{template}}</span>.</p>
      <p><strong>Último error recibido:</strong> {{error}}</p>
      <p>Hoja activa: <strong>{{sheetName}}</strong></p>
      <ol>
        <li>Confirma que el archivo exista en el proyecto de Apps Script y que no haya duplicados con nombres parecidos.</li>
        <li>Si trabajas con Clasp, ejecuta <code>clasp push</code> desde la carpeta donde está <code>.clasp.json</code>.</li>
        <li>Recarga la hoja de cálculo y vuelve a abrir el panel.</li>
      </ol>
    </div>
  </body>
</html>`;

function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🧪 Pruebas Sidebar')
      .addItem('Abrir panel', 'showSidebar')
      .addToUi();
  } catch (error) {
    if (!e) {
      console.warn(
        'No se pudo crear el menú del sidebar porque SpreadsheetApp.getUi() no está disponible en este contexto.',
        error,
      );
      return;
    }
    throw error;
  }
}

function showSidebar() {
  const ui = SpreadsheetApp.getUi();
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  let htmlOutput;
  let lastError;

  try {
    htmlOutput = HtmlService.createHtmlOutputFromFile(SIDEBAR_FILE);
  } catch (error) {
    console.warn(`No se pudo cargar ${SIDEBAR_FILE}.html`, error);
    lastError = error;
  }

  if (!htmlOutput) {
    const message = lastError && lastError.message ? lastError.message : 'Plantilla no encontrada';
    const html = FALLBACK_SIDEBAR_HTML
      .replace(/{{template}}/g, escapeHtml(`${SIDEBAR_FILE}.html`))
      .replace(/{{error}}/g, escapeHtml(message))
      .replace(/{{sheetName}}/g, escapeHtml(sheetName));
    htmlOutput = HtmlService.createHtmlOutput(html);
  }

  htmlOutput.setTitle(SIDEBAR_TITLE).setWidth(320);
  ui.showSidebar(htmlOutput);
}

function getActiveSheetName() {
  return SpreadsheetApp.getActiveSheet().getName();
}

function escapeHtml(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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

