/***** === SCHEMA & AUTO-FORMULAS (no destructivo) === *****/
/* Requiere que existan:
   - SHEET_MASTER / NICE_SHEET_MASTER
   - MASTER_COLUMNS (lista canónica de columnas)
   - Helpers: getSheet_, getHeaders_, idxMap_, norm_, getCurrentUserEmail_,
              rowToObject_, computeKpis_, recomputeKpisForRow_ (de KPIs.gs)
   - (Opcional) syncRowToSupabase_(rowIndex), syncRowToNoco_(rowIndex)
*/

function menuUpgradeSchema_(){
  const res = ensureMasterSchema_();
  const msg = `Columnas iniciales: ${res.before}\nAñadidas: ${res.added.join(', ') || '—'}\nTotal final: ${res.after}`;
  SpreadsheetApp.getUi().alert('Esquema verificado ✅\n\n' + msg);

  // Backfill (KPIs y defaults)
  backfillComputedColumns_();
  SpreadsheetApp.getUi().alert('Recalculo de KPIs completado ✅');
}

/** Asegura que existan todas las columnas de MASTER_COLUMNS en 📦 Cotizaciones_Master. */
function ensureMasterSchema_(){
  const sh = getSheet_(SHEET_MASTER) || getSheet_(NICE_SHEET_MASTER);
  if(!sh) throw new Error('No encuentro la hoja MASTER.');

  // Lee encabezados actuales
  const headers = getHeaders_(sh).map(h => norm_(h));
  const Hset = new Set(headers);
  const added = [];

  // Agrega faltantes al final (no destructivo)
  MASTER_COLUMNS.forEach(col => {
    const cname = norm_(col);
    if (!Hset.has(cname)){
      // añadir columna al final
      sh.insertColumnAfter(sh.getLastColumn() || 1);
      const newColIndex = sh.getLastColumn();
      sh.getRange(1, newColIndex).setValue(cname);
      added.push(cname);
      Hset.add(cname);
    }
  });

  // Pequeños defaults útiles si no existen
  const H = idxMap_(getHeaders_(sh));
  if (H['fecha_registro'] != null){
    const lr = sh.getLastRow();
    if (lr >= 2) {
      const rng = sh.getRange(2, H['fecha_registro']+1, lr-1, 1);
      const vals = rng.getValues();
      let changed = false;
      for (let i=0;i<vals.length;i++){
        if (!vals[i][0]) { vals[i][0] = new Date(); changed = true; }
      }
      if (changed) rng.setValues(vals);
    }
  }

  // ID UUID si no existe
  if (H['id'] != null){
    const lr = sh.getLastRow();
    if (lr >= 2) {
      const rng = sh.getRange(2, H['id']+1, lr-1, 1);
      const vals = rng.getValues();
      let changed = false;
      for (let i=0;i<vals.length;i++){
        if (!norm_(vals[i][0])) { vals[i][0] = newUUID_(); changed = true; }
      }
      if (changed) rng.setValues(vals);
    }
  }

  return { before: headers.length, added, after: getHeaders_(sh).length };
}

/** Recalcula KPIs para todas las filas (usa computeKpis_ que ya tienes en KPIs.gs) */
function backfillComputedColumns_(){
  const sh = getSheet_(SHEET_MASTER) || getSheet_(NICE_SHEET_MASTER);
  if(!sh) return;
  const headers = getHeaders_(sh);
  const H = idxMap_(headers);

  // ¿Qué columnas impactan KPIs?
  const KPI_INPUTS = ['costo_presupuestado','acumulado_coste_real','acumulado_avance_valorizado'];
  KPI_INPUTS.forEach(name=>{
    if (H[name]==null) Logger.log('⚠ Falta columna de entrada KPI: ' + name);
  });

  const lr = sh.getLastRow();
  if (lr < 2) return;

  const data = sh.getRange(2,1,lr-1,headers.length).getValues();
  for (let r = 0; r < data.length; r++){
    const rowIndex = r + 2;
    try{
      const obj = rowToObject_(headers, data[r]);
      const k = computeKpis_(obj); // retorna {cpi, spi, eac, etc}
      if (H['cpi']!=null) sh.getRange(rowIndex, H['cpi']+1).setValue(k.cpi);
      if (H['spi']!=null) sh.getRange(rowIndex, H['spi']+1).setValue(k.spi);
      if (H['eac']!=null) sh.getRange(rowIndex, H['eac']+1).setValue(k.eac);
      if (H['etc']!=null) sh.getRange(rowIndex, H['etc']+1).setValue(k.etc);
    }catch(e){
      Logger.log('Fila '+rowIndex+': '+e);
    }
  }
}

/** onEdit: si cambian entradas clave, recalcula KPIs para esa fila y sincroniza */
function onEdit(e){
  try{
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (!isMasterSheet_(sh)) return;

    const headers = getHeaders_(sh);
    const H = idxMap_(headers);
    const editedCol = e.range.getColumn()-1;

    // Columnas que gatillan KPIs
    const KPI_INPUTS = new Set(['costo_presupuestado','acumulado_coste_real','acumulado_avance_valorizado','monto_ofertado']);
    const editedHeader = headers[editedCol] ? String(headers[editedCol]) : '';
    if (KPI_INPUTS.has(editedHeader)){
      const row = e.range.getRow();
      if (row >= 2){
        const k = recomputeKpisForRow_(row); // ya la tienes en KPIs.gs
        // Sincroniza si tienes implementado
        safeSyncRow_(row);
      }
    }

    // Si se crea nueva fila o se edita 'id' vacío, auto ID y fecha_registro
    if (e.value && e.range.getRow()>=2){
      const row = e.range.getRow();
      if (H['id']!=null){
        const idCell = sh.getRange(row, H['id']+1);
        if (!norm_(idCell.getValue())) idCell.setValue(newUUID_());
      }
      if (H['fecha_registro']!=null){
        const fCell = sh.getRange(row, H['fecha_registro']+1);
        if (!fCell.getValue()) fCell.setValue(new Date());
      }
    }
  }catch(err){
    Logger.log('onEdit error: ' + err);
  }
}

/** Wrapper: si existen funciones de sync, las usa sin romper si no existen */
function safeSyncRow_(rowIndex){
  try{
    if (typeof syncRowToSupabase_ === 'function'){
      syncRowToSupabase_(rowIndex);
    }
  }catch(e){ Logger.log('sync supabase opcional: '+e); }
  try{
    if (typeof syncRowToNoco_ === 'function'){
      syncRowToNoco_(rowIndex);
    }
  }catch(e){ Logger.log('sync noco opcional: '+e); }
}

/** Util: objeto fila (por nombre de encabezados) */
function rowToObject_(headers, rowVals){
  const o = {};
  headers.forEach((h,i)=>{ o[String(h)] = rowVals[i]; });
  return o;
}

/** KPIs básicos tipo EVM, robustos a vacíos */
function computeKpis_(o){
  const toN = v => (v==='' || v==null) ? NaN : Number(v);
  const EV  = toN(o['acumulado_avance_valorizado']); // Earned Value
  const AC  = toN(o['acumulado_coste_real']);        // Actual Cost
  const BAC = toN(o['costo_presupuestado']);         // Presupuesto a la conclusión

  let cpi = (isFinite(EV) && isFinite(AC) && AC>0) ? EV/AC : '';
  let spi = (isFinite(EV) && isFinite(BAC) && BAC>0) ? EV/BAC : ''; // Aproximación (PV no disponible)
  let eac = (isFinite(cpi) && cpi>0 && isFinite(BAC)) ? (BAC/cpi) : '';
  let etc = (isFinite(eac) && isFinite(AC)) ? (eac - AC) : '';

  // Redondeo amable
  const rnd = (x)=> (x===''? '': Math.round(x*100)/100);
  return {
    cpi: rnd(cpi),
    spi: rnd(spi),
    eac: rnd(eac),
    etc: rnd(etc)
  };
}

/** UUID v4 simple */
function newUUID_(){
  const s = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
  return s.replace(/[xy]/g, c=>{
    const r = Math.random()*16|0, v = c=='x'? r : (r&0x3|0x8);
    return v.toString(16);
  });
}
