/***** UTILIDADES *****/
function getSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.getSheetByName(NAME_MAP.get(name)) || null;
}
function getOrCreateSheet_(name){
  return getSheet_(name) || SpreadsheetApp.getActive().insertSheet(name);
}
function getHeaders_(sheet){
  const lc = sheet.getLastColumn();
  return lc ? sheet.getRange(1,1,1,lc).getValues()[0] : [];
}
function getData_(sheet){
  const lr=sheet.getLastRow(), lc=sheet.getLastColumn();
  return (lr>=2&&lc)? sheet.getRange(2,1,lr-1,lc).getValues():[];
}
function idxMap_(headers){ const m={}; headers.forEach((h,i)=>m[String(h)]=i); return m; }
function norm_(s){ return String(s||'').trim(); }
function rgbToHex_(o){ const t=v=>Math.max(0,Math.min(255,Math.round((v||0)*255))); return '#'+[o.r,o.g,o.b].map(v=>t(v).toString(16).padStart(2,'0')).join(''); }
function getCurrentUserEmail_(){ return (Session.getEffectiveUser().getEmail()||'').toLowerCase(); }
function isMasterSheet_(sh){ const n=sh.getName(); return (n===SHEET_MASTER || n===NAME_MAP.get(SHEET_MASTER)); }

/***** Helpers catálogos *****/
function ensureCatalogSheets_(){
  // Cliente
  let sh = getOrCreateSheet_(SHEET_CAT_CLIENTE);
  if (sh.getLastRow()===0) sh.appendRow(['Cliente']);

  // Zona
  sh = getOrCreateSheet_(SHEET_CAT_ZONA);
  if (sh.getLastRow()===0) sh.appendRow(['Zona de Trabajo']);

  // Solicitante
  sh = getOrCreateSheet_(SHEET_CAT_SOLI);
  if (sh.getLastRow()===0) sh.appendRow(['Nombre','Correo','Teléfono']);

  // Personal EKA
  sh = getOrCreateSheet_(SHEET_CAT_PERSONAL);
  if (sh.getLastRow()===0) sh.appendRow(['Nombre','Correo','Teléfono','Rol']);

  // Industria
  sh = getOrCreateSheet_(SHEET_CAT_INDUSTRIA);
  if (sh.getLastRow()===0) sh.appendRow(['Industria']);
}

function seedRow_(sheetName, arr){
  const sh = getOrCreateSheet_(sheetName);
  sh.appendRow(arr);
  return true;
}

/***** === VALIDACIONES (listas desplegables) === *****/
function menuSetupDropdowns_(){
  ensureCatalogSheets_();

  const ms = getOrCreateSheet_(SHEET_MASTER);
  const h  = getHeaders_(ms);
  const H  = idxMap_(h);
  const maxRows = ms.getMaxRows();

  function setListValidation_(colHeader, rangeSheetName){
    const col = H[colHeader];
    if (col == null) return; // la columna no existe en el master

    const cat = getOrCreateSheet_(rangeSheetName);
    const last = Math.max(cat.getLastRow(),2);
    const src = cat.getRange(2,1,last-1,1); // una columna

    const rng = ms.getRange(2, col+1, maxRows-1, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(src, true) // ✅ correcto para listas desde rango
      .setAllowInvalid(false)
      .build();
    rng.setDataValidation(rule);
  }

  setListValidation_('Cliente', SHEET_CAT_CLIENTE);
  setListValidation_('Zona de Trabajo', SHEET_CAT_ZONA);
  setListValidation_('Solicitante', SHEET_CAT_SOLI);
  setListValidation_('Industria', SHEET_CAT_INDUSTRIA);
  setListValidation_('Responsable Técnico', SHEET_CAT_PERSONAL);
  setListValidation_('Responsable Económico', SHEET_CAT_PERSONAL);

  SpreadsheetApp.getUi().alert('Listas desplegables configuradas ✅');
}

/***** === Carpeta por cotización === *****/
function getActiveCotizacionContext_(){
  const sh = SpreadsheetApp.getActive().getActiveSheet();
  if(!isMasterSheet_(sh)) return {ok:false, reason:'Debes estar en la hoja Cotizaciones_Master.'};
  const H = idxMap_(getHeaders_(sh));
  const key = H['Cotización']!=null ? 'Cotización' : (H['cotizacion']!=null ? 'cotizacion' : null);
  if (key==null) return {ok:false, reason:'No encuentro la columna "Cotización".'};
  const rng = SpreadsheetApp.getActive().getActiveRange();
  if(!rng || rng.getRow()<2 || rng.getColumn()!==H[key]+1) return {ok:false, reason:'Ubica la celda activa en la columna "Cotización".'};
  const row = rng.getRow();
  const rich = sh.getRange(row, H[key]+1).getRichTextValue();
  const val = norm_(rich ? rich.getText() : sh.getRange(row, H[key]+1).getDisplayValue());
  const url = rich ? rich.getLinkUrl() : '';
  let folderId = '';
  if(url && /drive\.google\.com\/drive\/folders\//.test(url)) folderId = url.split('/folders/')[1].split(/[/?#]/)[0];
  return {ok:true, cotizacion:val, folderUrl:url||'', folderId, row, col:H[key]+1};
}

function createProjectFolders_(name){
  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const folder = parent.createFolder(name);
  [
    '00_Entrada de Licitación (invitación, TDR, anexos)',
    '05_Alcance (PGA, Requisitos, EDT/WBS, diccionario)',
    '06_Cronograma (MSP nivel 7, Curva S, línea base)',
    '07_Costos (estimaciones, presupuesto, supuestos)',
    '08_Calidad (Plan de Calidad, formatos de control)',
    '09_Recursos (RRHH, RACI, perfiles críticos)',
    '10_Comunicaciones (matriz, formatos de reporte)',
    '11_Riesgos (registro, análisis cuali/cuant, respuesta)',
    '12_Adquisiciones (Plan)',
    '12_Adquisiciones/Pre-OC (cotizaciones de materiales)',
    '12_Adquisiciones/Subcontratos (pre-adjudicación)',
    '12_Adquisiciones/Requerimientos (prediseños)',
    '13_Interesados (registro, estrategia)',
    'Control Documentario (índices, dossier preliminar)'
  ].forEach(n=>folder.createFolder(n));
  return { id: folder.getId(), url: folder.getUrl() };
}

/***** === SERVIDOR: usado por Sidebar (google.script.run) === *****/
function include_(filename){ return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function getServerInfo_(){
  return {
    user: getCurrentUserEmail_(),
    sb: {url: SB_URL, table: SB_TABLE, ok: !!SB_URL},
    noco: {url: NOCO_URL, project: NOCO_PROJECT, table: NOCO_TABLE, ok: !!NOCO_URL}
  };
}

function serverGetInit_(){ return getServerInfo_(); }
function serverPing_(){ return 'pong'; }

function serverSetupDropdowns_(){ menuSetupDropdowns_(); return true; }

function serverCreateFolderForActive_(){
  const ctx = getActiveCotizacionContext_();
  if(!ctx.ok) throw new Error(ctx.reason);
  if(!ctx.cotizacion) throw new Error('La celda "Cotización" está vacía.');
  const made = createProjectFolders_(ctx.cotizacion);
  // enlazar la celda con link
  const sh = SpreadsheetApp.getActive().getActiveSheet();
  const rich = SpreadsheetApp.newRichTextValue().setText(ctx.cotizacion).setLinkUrl(made.url).build();
  sh.getRange(ctx.row, ctx.col).setRichTextValue(rich);
  return made;
}

function serverRecomputeKpisAll_(){
  if (typeof backfillComputedColumns_ === 'function') backfillComputedColumns_();
  return true;
}

function serverRebuildPermissions_(){
  if (typeof rebuildPermissions_ === 'function') { rebuildPermissions_(); return true; }
  return false;
}

/***** === Altas rápidas de Catálogos === *****/
function addCliente(nombre){ ensureCatalogSheets_(); seedRow_(SHEET_CAT_CLIENTE, [nombre]); return true; }
function addZona(zona){ ensureCatalogSheets_(); seedRow_(SHEET_CAT_ZONA, [zona]); return true; }
function addIndustria(ind){ ensureCatalogSheets_(); seedRow_(SHEET_CAT_INDUSTRIA, [ind]); return true; }

// Solicitante con correo/teléfono
function addSolicitante(nombre, correo, tel){
  ensureCatalogSheets_();
  seedRow_(SHEET_CAT_SOLI, [nombre||'', correo||'', tel||'']);
  return true;
}

// Personal EKA con correo/teléfono/rol
function addPersona(nombre, correo, tel, rol){
  ensureCatalogSheets_();
  seedRow_(SHEET_CAT_PERSONAL, [nombre||'', correo||'', tel||'', rol||'']);
  return true;
}

/***** === Sidebar === *****/
function showSidebar(){
  // Usar template para evaluar <?!= include_('PanelJS'); ?>
  const t = HtmlService.createTemplateFromFile('Sidebar');
  const html = t.evaluate()
    .setTitle('Cotizaciones')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

/***** === Menú === *****/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('🚀 Cotizaciones')
    .addItem('📂 Abrir Sidebar','showSidebar')
    .addSeparator()
    .addItem('🧰 Configurar listas desplegables','menuSetupDropdowns_')
    .addItem('📁 Crear carpeta para la fila activa','serverCreateFolderForActive_')
    .addItem('🔄 Sincronizar fila activa (KPIs/Supabase/NocoDB)','serverPing_') // placeholder si deseas
    .addSeparator()
    .addItem('👥 Usuarios: actualizar permisos','serverRebuildPermissions_')
    .addItem('📈 KPIs: recalcular todo','serverRecomputeKpisAll_')
    .addToUi();
}


