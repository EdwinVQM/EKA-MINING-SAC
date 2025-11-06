const SIDEBAR_TITLE = 'Panel de pruebas';
const SIDEBAR_TEMPLATES = ['Sidebar'];
/**
 * Plantilla HTML utilizada cuando no se encuentra ningún archivo válido.
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
      <p class="error">No se pudo cargar la plantilla <span class="template">{{lastTemplate}}</span>.</p>
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

function showSidebar(templateName) {
  const ui = SpreadsheetApp.getUi();
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  const attemptedTemplates = [];
  const candidates = getTemplateCandidates(templateName);
  let htmlOutput;
  let lastError;
  let lastTemplate;

  for (const candidate of candidates) {
    attemptedTemplates.push(`${candidate}.html`);
    try {
      const template = HtmlService.createTemplateFromFile(candidate);
      template.sheetName = sheetName;
      htmlOutput = template.evaluate();
      lastTemplate = `${candidate}.html`;
      break;
    } catch (error) {
      console.warn(`No se pudo cargar ${candidate}.html`, error);
      lastError = error;
      lastTemplate = `${candidate}.html`;
    }
  }

  if (!htmlOutput) {
    console.error('No se pudo renderizar ninguna plantilla del sidebar.', lastError);
    const message = lastError && lastError.message ? lastError.message : String(lastError || 'Plantilla no encontrada');
    const templateListHtml = attemptedTemplates
      .map((name) => `<code>${escapeHtml(name)}</code>`)
      .join(', ');
    const lastTemplateHtml = lastTemplate
      ? escapeHtml(lastTemplate)
      : '<em>ninguna</em>';
    const html = FALLBACK_SIDEBAR_HTML
      .replace(/{{error}}/g, escapeHtml(message))
      .replace(/{{sheetName}}/g, escapeHtml(sheetName))
      .replace(/{{templates}}/g, templateListHtml || '<em>sin registros</em>')
      .replace(/{{lastTemplate}}/g, lastTemplateHtml);
    htmlOutput = HtmlService.createHtmlOutput(html);
  }

  htmlOutput.setTitle(SIDEBAR_TITLE).setWidth(320);
  ui.showSidebar(htmlOutput);
}

function getTemplateCandidates(templateName) {
  const providedTemplates = Array.isArray(templateName)
    ? templateName
    : [templateName].filter((name) => name !== undefined && name !== null);

  const normalized = providedTemplates
    .map(normalizeTemplateName)
    .filter(Boolean);

  const baseCandidates = normalized.length ? normalized : SIDEBAR_TEMPLATES;
  const expanded = [];

  for (const name of baseCandidates) {
    if (!name) {
      continue;
    }
    const variants = new Set([name, name.toLowerCase(), name.toUpperCase()]);
    variants.forEach((value) => {
      const normalizedValue = normalizeTemplateName(value);
      if (normalizedValue) {
        expanded.push(normalizedValue);
      }
    });
  }

  return Array.from(new Set(expanded));
}

function normalizeTemplateName(name) {
  if (typeof name !== 'string') {
    return '';
  }
  return name
    .trim()
    .replace(/\.html?$/i, '')
    .trim();
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

