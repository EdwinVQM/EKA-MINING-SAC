/********** MAPEO FILA <-> OBJETO **********/
function rowToObject_(headers, row){
  const o = {};
  headers.forEach((h,i)=>{
    let v = row[i];
    // normaliza fechas a ISO en campos que suenan a fecha
    if (String(h).toLowerCase().startsWith('fecha') || ['mes_inicio_forecast','mes_fin_forecast'].includes(h)){
      v = parseDateToISO_(v);
    }
    o[h] = v === '' ? null : v;
  });
  // id autogenerado si viene vacío
  if (!o.id) o.id = ensureUuid_();
  // probabilidad num
  if (o.probabilidad!=null) o.probabilidad = Number(o.probabilidad)||0;
  // montos num
  ['monto_estimado','monto_ofertado','valor_contrato','costo_presupuestado','acumulado_coste_real','acumulado_avance_valorizado','eac','etc','cpi','spi','kpi_score']
    .forEach(k=>{ if(o[k]!=null) o[k]=Number(o[k])||0; });
  return o;
}

function objectToRow_(headers, obj){
  return headers.map(h => h in obj ? obj[h] : '');
}

/********** SUPABASE **********/
function supabaseUpsert_(obj){
  if(!SB_URL || !SB_KEY) return {ok:false, reason:'Supabase no configurado'};
  const url = `${SB_URL}/rest/v1/${encodeURIComponent(SB_TABLE)}`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify([obj]),
    headers: {
      'apikey': SB_KEY,
      'Authorization': `Bearer ${SB_KEY}`,
      'Prefer': 'resolution=merge-duplicates' // upsert
    },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code>=200 && code<300) return {ok:true};
  return {ok:false, reason:`SB ${code}: ${res.getContentText()}`};
}

/********** NOCO DB **********/
function nocoUpsert_(obj){
  if(!NOCO_URL || !NOCO_TOKEN || !NOCO_PROJECT) return {ok:false, reason:'NocoDB no configurado'};
  const base = `${NOCO_URL}/api/v2/tables/${encodeURIComponent(NOCO_PROJECT)}/${encodeURIComponent(NOCO_TABLE)}`;
  // Usaremos id como unique. Intentar obtener y luego patch; si no existe, crear.
  try{
    const getUrl = `${base}/rows?where=(id,eq,${encodeURIComponent(obj.id)})`;
    const getRes = UrlFetchApp.fetch(getUrl, {
      method:'get',
      headers:{'xc-token': NOCO_TOKEN},
      muteHttpExceptions:true
    });
    if (getRes.getResponseCode()===200){
      const data = JSON.parse(getRes.getContentText());
      if (Array.isArray(data.list) && data.list.length>0){
        // update
        const rowId = data.list[0]['Id'] || data.list[0]['id'] || data.list[0]['_id'] || data.list[0]['__id'] || data.list[0]['Id'];
        const patchUrl = `${base}/rows/${encodeURIComponent(rowId)}`;
        const patchRes = UrlFetchApp.fetch(patchUrl, {
          method:'patch',
          headers:{'xc-token': NOCO_TOKEN,'Content-Type':'application/json'},
          payload: JSON.stringify(obj),
          muteHttpExceptions:true
        });
        const code = patchRes.getResponseCode();
        if (code>=200 && code<300) return {ok:true};
        return {ok:false, reason:`Noco PATCH ${code}: ${patchRes.getContentText()}`};
      }
    }
    // create
    const postRes = UrlFetchApp.fetch(`${base}/rows`, {
      method:'post',
      headers:{'xc-token': NOCO_TOKEN,'Content-Type':'application/json'},
      payload: JSON.stringify(obj),
      muteHttpExceptions:true
    });
    const pcode = postRes.getResponseCode();
    if (pcode>=200 && pcode<300) return {ok:true};
    return {ok:false, reason:`Noco POST ${pcode}: ${postRes.getContentText()}`};
  }catch(e){
    return {ok:false, reason: String(e)};
  }
}

/********** SYNC UNIFICADO **********/
function syncRow_(obj){
  const r1 = supabaseUpsert_(obj);
  const r2 = nocoUpsert_(obj);
  const errs = [];
  if(!r1.ok) errs.push(r1.reason);
  if(!r2.ok) errs.push(r2.reason);
  return errs.length? {ok:false, reason:errs.join(' | ')} : {ok:true};
}

/********** KPI AUTO-CÁLCULO **********/
function computeKpis_(o){
  const ev  = Number(o.acumulado_avance_valorizado) || 0; // Earned Value
  const ac  = Number(o.acumulado_coste_real) || 0;        // Actual Cost
  const pv  = Number(o.costo_presupuestado) || 0;         // Planned Value (base presup.)
  const cpi = ac>0 ? ev/ac : null;
  const spi = pv>0 ? ev/pv : null;
  const bac = Number(o.valor_contrato || o.monto_ofertado || 0); // valor total contrato/ofertado
  const eac = (cpi && cpi>0) ? bac / cpi : null;
  const etc = (eac!=null) ? eac - ac : null;
  return {
    cpi: cpi==null? null : Number(cpi.toFixed(3)),
    spi: spi==null? null : Number(spi.toFixed(3)),
    eac: eac==null? null : Number(eac.toFixed(2)),
    etc: etc==null? null : Number(etc.toFixed(2)),
  };
}

/********** TRIGGERS **********/
function onEdit(e){
  try{
    const sh = e.range.getSheet();
    if(!isMasterSheet_(sh)) return;

    const headers = getHeaders_(sh);
    const H = idxMap_(headers);
    if (e.range.getRow() < 2) return;
    const rowIdx = e.range.getRow();
    const rowVals = sh.getRange(rowIdx,1,1,headers.length).getValues()[0];
    let obj = rowToObject_(headers, rowVals);

    // autocompletar id y fecha_registro al crear
    if(!obj.id) { obj.id = ensureUuid_(); sh.getRange(rowIdx, H['id']+1).setValue(obj.id); }
    if(!obj.fecha_registro){ const now = new Date(); sh.getRange(rowIdx, H['fecha_registro']+1).setValue(now); obj.fecha_registro = now.toISOString(); }

    // KPIs
    const k = computeKpis_(obj);
    ['cpi','spi','eac','etc'].forEach(key=>{
      if (H[key]!=null){
        sh.getRange(rowIdx, H[key]+1).setValue(k[key]);
        obj[key] = k[key];
      }
    });

    // Sync
    const sync = syncRow_(obj);
    if(!sync.ok) Logger.log(`SYNC ERROR fila ${rowIdx}: ${sync.reason}`);
  }catch(err){
    Logger.log('onEdit error: '+err);
  }
}

/** Re-sincroniza explícitamente la fila activa (desde menú / sidebar) */
function resyncActiveRow_(){
  const sh = getSheet_(SHEET_MASTER) || getSheet_(NICE_SHEET_MASTER);
  const rng = sh.getActiveRange();
  if(!rng || rng.getRow()<2) { SpreadsheetApp.getUi().alert('Selecciona una fila válida (>=2).'); return; }
  const headers = getHeaders_(sh);
  const rowVals = sh.getRange(rng.getRow(),1,1,headers.length).getValues()[0];
  let obj = rowToObject_(headers, rowVals);
  const k = computeKpis_(obj);
  ['cpi','spi','eac','etc'].forEach(key=>{
    const H = idxMap_(headers);
    if (H[key]!=null){ sh.getRange(rng.getRow(), H[key]+1).setValue(k[key]); obj[key] = k[key]; }
  });
  const sync = syncRow_(obj);
  SpreadsheetApp.getUi().alert(sync.ok? 'Resync OK ✅' : `Error: ${sync.reason}`);
}

/** Re-sincroniza todas las filas (cautela si hay muchas) */
function resyncAll_(){
  const sh = getSheet_(SHEET_MASTER) || getSheet_(NICE_SHEET_MASTER);
  const headers = getHeaders_(sh);
  const H = idxMap_(headers);
  const data = getData_(sh);
  let ok=0, err=0;
  data.forEach((row,i)=>{
    let obj = rowToObject_(headers,row);
    // asegurar id/fecha_registro
    if(!obj.id){ obj.id=ensureUuid_(); sh.getRange(i+2, H['id']+1).setValue(obj.id); }
    if(!obj.fecha_registro){ const now=new Date(); sh.getRange(i+2, H['fecha_registro']+1).setValue(now); obj.fecha_registro = now.toISOString(); }
    // KPIs
    const k = computeKpis_(obj);
    ['cpi','spi','eac','etc'].forEach(key=>{
      if (H[key]!=null){
        sh.getRange(i+2, H[key]+1).setValue(k[key]);
        obj[key] = k[key];
      }
    });
    const r = syncRow_(obj);
    r.ok ? ok++ : err++;
  });
  SpreadsheetApp.getUi().alert(`Resync finalizado. OK: ${ok} / ERR: ${err}`);
}
