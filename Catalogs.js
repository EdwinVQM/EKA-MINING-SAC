/***** === CATALOGS.gs (listas y catálogos) === *****/

/** Helpers mínimos (usa los que ya tienes si existen) */
function _getSS_(){ return SpreadsheetApp.getActive(); }
function _getOrCreateSheetByName_(name){
  const ss = _getSS_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function _getHeaders_(sheet){
  const lc = sheet.getLastColumn();
  return lc ? sheet.getRange(1,1,1,lc).getValues()[0] : [];
}
function _idxMap_(headers){ const m={}; headers.forEach((h,i)=>m[String(h)]=i); return m; }
function _norm_(s){ return String(s||'').trim(); }

/** === Crear las hojas de catálogo si no existen, con cabecera en A1 === */
function ensureCatalogSheet_(name, header){
  const sh = _getOrCreateSheetByName_(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1).setValue(header);
  return sh;
}

/** === Agregar único valor (case-insensitive) a una hoja catálogo === */
function pushUniqueToCatalog_(sheetName, value){
  const v = _norm_(value);
  if (!v) return {ok:false, reason:'Valor vacío'};
  const sh = ensureCatalogSheet_(sheetName, sheetName);
  const data = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),1).getValues().map(r=>_norm_(r[0]).toLowerCase()).filter(Boolean);
  if (data.indexOf(v.toLowerCase()) === -1){
    sh.appendRow([v]);
  }
  return {ok:true, value:v};
}

/** === Endpoints para el Sidebar: Altas rápidas de catálogo === */
function addCliente_(nombre){ return pushUniqueToCatalog_('Cliente', nombre); }
function addZonaTrabajo_(nombre){ return pushUniqueToCatalog_('Zona de Trabajo', nombre); }
function addSolicitante_(nombre){ return pushUniqueToCatalog_('Solicitante', nombre); }
function addPersonalEKA_(nombre){ return pushUniqueToCatalog_('Personal EKA', nombre); }
function addIndustria_(nombre){ return pushUniqueToCatalog_('Industria', nombre); }

/**
 * === Configurar listas desplegables en "Cotizaciones_Master"
 * Aplica validación por rango (A2:A) de cada hoja de catálogo a su columna en Master.
 * NO borra datos. Requiere que la fila 1 de Master tenga encabezados.
 */
function setupDropdowns_(){
  const ss = _getSS_();
  const master = ss.getSheetByName('Cotizaciones_Master') || ss.getSheetByName('📦 Cotizaciones_Master');
  if (!master) throw new Error('No encuentro la hoja "Cotizaciones_Master".');

  // Asegurar hojas catálogo
  ensureCatalogSheet_('Cliente','Cliente');
  ensureCatalogSheet_('Zona de Trabajo','Zona de Trabajo');
  ensureCatalogSheet_('Solicitante','Solicitante');
  ensureCatalogSheet_('Personal EKA','Personal EKA');
  ensureCatalogSheet_('Industria','Industria');

  const headers = _getHeaders_(master);
  if (!headers.length) throw new Error('La hoja Master no tiene encabezados en la fila 1.');
  const H = _idxMap_(headers);

  // Mapeo: encabezado Master -> hoja catálogo
  const mapping = [
    {header:'Cliente',               sheet:'Cliente'},
    {header:'Zona de Trabajo',       sheet:'Zona de Trabajo'},
    {header:'Solicitante',           sheet:'Solicitante'},
    {header:'Responsable Técnico',   sheet:'Personal EKA'},
    {header:'Responsable Económico', sheet:'Personal EKA'},
    {header:'Industria',             sheet:'Industria'}
  ];

  // Rango de filas de datos en Master
  const lastRow = Math.max(master.getLastRow(), 2);
  const dataRows = lastRow - 1; // sin cabecera
  if (dataRows <= 0){
    // Igual configuramos validación en un rango "grande" para futuras filas (hasta 2000 filas)
    const maxRows = 2000;
    mapping.forEach(m=>{
      if (H[m.header] == null) return;
      const shCat = ss.getSheetByName(m.sheet);
      const catRange = shCat.getRange('A2:A');
      const col = H[m.header] + 1;
      const rng = master.getRange(2, col, maxRows, 1);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(catRange, true)   // <- FIX: requireValueInRange
        .setAllowInvalid(false)
        .build();
      rng.setDataValidation(rule);
    });
    return 'Listas configuradas (rango anticipado).';
  }

  // Configurar validación para las filas actualmente existentes
  mapping.forEach(m=>{
    if (H[m.header] == null) return;
    const shCat = ss.getSheetByName(m.sheet);
    const catRange = shCat.getRange('A2:A'); // lista
    const col = H[m.header] + 1;
    const rng = master.getRange(2, col, dataRows, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(catRange, true)     // <- FIX
      .setAllowInvalid(false)
      .build();
    rng.setDataValidation(rule);
  });

  return 'Listas desplegables configuradas en Cotizaciones_Master.';
}

/** ======= Nombres de hojas (ajústalos si difieren) ======= */
const SHEET_CAT_CLIE   = 'Cliente';
const SHEET_CAT_ZONAS  = 'Zona de Trabajo';
const SHEET_CAT_SOLI   = 'Solicitante';
const SHEET_CAT_INDU   = 'Industria';
const SHEET_CAT_PERS   = 'Personal EKA';

/** ======= Utilitarios ======= */
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeader_(sh, headers) {
  const first = sh.getRange(1,1,1,headers.length).getValues()[0];
  const needs = first.join('') === '' || first.length < headers.length;
  if (needs) sh.getRange(1,1,1,headers.length).setValues([headers]);
}

function seedRow_(sheetName, headers, values) {
  const sh = getOrCreateSheet_(sheetName);
  ensureHeader_(sh, headers);
  sh.appendRow(values);
  return { ok: true, sheet: sheetName, row: sh.getLastRow() };
}

/** ======= Altas rápidas ======= */
function addCliente(nombre) {
  nombre = (nombre || '').toString().trim();
  if (!nombre) throw new Error('Ingresa un nombre de cliente.');
  return seedRow_(SHEET_CAT_CLIE, ['Cliente'], [nombre]);
}

function addZona(cliente, unidad) {
  if (!(cliente||'').trim()) throw new Error('Cliente requerido.');
  if (!(unidad||'').trim())  throw new Error('Zona/Unidad requerida.');
  return seedRow_(SHEET_CAT_ZONAS, ['Cliente','Unidad','Cod','Obs'], [cliente, unidad, '', '']);
}

function addSolicitante(nombre, correo, telefono) {
  if (!(nombre||'').trim()) throw new Error('Nombre requerido.');
  return seedRow_(SHEET_CAT_SOLI, ['Nombre','Correo','Teléfono'], [nombre||'', correo||'', telefono||'']);
}

function addIndustria(nombre) {
  if (!(nombre||'').trim()) throw new Error('Industria requerida.');
  return seedRow_(SHEET_CAT_INDU, ['Industria'], [nombre]);
}

function addPersonalEKA(nombre, correo, telefono, rol) {
  if (!(nombre||'').trim()) throw new Error('Nombre requerido.');
  return seedRow_(SHEET_CAT_PERS, ['Nombre','Correo','Teléfono','Rol'], [nombre||'', correo||'', telefono||'', rol||'']);
}

/** ======= Ping / Carga inicial para "Probar conexión" ======= */
function getInitData() {
  const ss = SpreadsheetApp.getActive();
  const data = [SHEET_CAT_CLIE, SHEET_CAT_ZONAS, SHEET_CAT_SOLI, SHEET_CAT_INDU, SHEET_CAT_PERS]
    .map(n => {
      const sh = getOrCreateSheet_(n);
      return { sheet: n, rows: sh.getLastRow(), cols: sh.getLastColumn() };
    });
  return { ok: true, spreadsheet: ss.getName(), data, ts: new Date() };
}
