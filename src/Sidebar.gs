const SIDEBAR_TITLE = 'Panel de pruebas';
const SIDEBAR_TEMPLATES = ['SidebarMain', 'Sidebar'];
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
      <p class="error">No se pudo cargar la plantilla <code>{{lastTemplate}}</code>.</p>
      <p>Se intentó abrir: {{templates}}.</p>
      <p>Verifica que alguno de esos archivos exista en el proyecto de Apps Script y vuelve a ejecutar el menú.</p>
      <p><strong>Último error recibido:</strong> {{error}}</p>
      <p>Hoja activa: <strong>{{sheetName}}</strong></p>
      <ol>
        <li>Confirma que el repositorio contenga los archivos HTML listados arriba.</li>
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

function showSidebar(templateName) {
  const ui = SpreadsheetApp.getUi();
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  const attemptedTemplates = [];
  const templates = Array.isArray(templateName)
    ? templateName
    : [templateName].filter(Boolean);
  const candidates = templates.length ? templates : SIDEBAR_TEMPLATES;
  let htmlOutput;
  let lastError;
  let lastTemplate;

  for (const candidate of candidates) {
    attemptedTemplates.push(candidate);
    try {
      const template = HtmlService.createTemplateFromFile(candidate);
      template.sheetName = sheetName;
      htmlOutput = template.evaluate();
      lastTemplate = candidate;
      break;
    } catch (error) {
      console.warn(`No se pudo cargar ${candidate}.html`, error);
      lastError = error;
      lastTemplate = candidate;
    }
  }

  if (!htmlOutput) {
    console.error('No se pudo renderizar ninguna plantilla del sidebar.', lastError);
    const message = lastError && lastError.message ? lastError.message : String(lastError || 'Plantilla no encontrada');
    const templateListHtml = attemptedTemplates
      .map((name) => `<code>${name}.html</code>`)
      .join(', ');
    const lastTemplateHtml = lastTemplate
      ? `<code>${lastTemplate}.html</code>`
      : '<em>ninguna</em>';
    const html = FALLBACK_SIDEBAR_HTML
      .replace(/{{error}}/g, message)
      .replace(/{{sheetName}}/g, sheetName)
      .replace(/{{templates}}/g, templateListHtml || '<em>sin registros</em>')
      .replace(/{{lastTemplate}}/g, lastTemplateHtml);
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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
