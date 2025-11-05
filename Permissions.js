/** ===================== PERMISSIONS.GS ===================== **
 * Control de roles y protecciones de columnas en Cotizaciones_Master
 * Roles: admin, editor, viewer
 * - admin: puede editar TODO, incluidas columnas no editables
 * - editor: puede editar solo columnas marcadas como 'sí' en 🧱 Columnas
 * - viewer: sin efecto aquí (control por compartición de archivo), pero queda documentado
 */

/***** Helpers de Usuarios *****/
function ensureUsersSheet_(){
  const sh = getOrCreateSheet_(SHEET_USERS) || getOrCreateSheet_(NICE_SHEET_USERS);
  if (sh.getLastRow() === 0){
    sh.getRange(1,1,1,3).setValues([['email','rol','nota']]);
  } else {
    // Normaliza header
    const hdr = getHeaders_(sh);
    const need = ['email','rol','nota'];
    if (hdr.length < 3 || hdr[0] !== 'email' || hdr[1] !== 'rol'){
      const merged = ['email','rol','nota'];
      sh.clearContents();
      sh.getRange(1,1,1,merged.length).setValues([merged]);
    }
  }
  return sh;
}

function getUsers_(){
  const sh = ensureUsersSheet_();
  const data = getData_(sh);
  const users = data
    .map(r => ({email: String(r[0]||'').trim().toLowerCase(), rol: String(r[1]||'').trim().toLowerCase()}))
    .filter(u => u.email);
  // agrega ADMIN_EMAIL como admin siempre
  if (ADMIN_EMAIL) {
    const adminExists = users.some(u => u.email === ADMIN_EMAIL.toLowerCase());
    if (!adminExists) users.push({email: ADMIN_EMAIL.toLowerCase(), rol: 'admin'});
  }
  return users;
}

function getAdmins_(){
  return getUsers_().filter(u => u.rol === 'admin').map(u => u.email);
}

function getEditors_(){
  return getUsers_().filter(u => u.rol === 'editor').map(u => u.email);
}

function isAdmin_(email){
  const e = String(email||'').toLowerCase();
  if (!e) return false;
  if (e === String(ADMIN_EMAIL||'').toLowerCase()) return true;
  return getAdmins_().includes(e);
}

/***** Hoja 🧱 Columnas -> columnas editables *****/
function getEditableColumns_(){
  const sh = getOrCreateSheet_(SHEET_COLS) || getOrCreateSheet_(NICE_SHEET_COLS);
  if (!sh) return new Set(DEFAULT_EDITABLE);
  const H = getHeaders_(sh);
  const map = idxMap_(H);
  if (map['columna']==null || map['editable']==null) return new Set(DEFAULT_EDITABLE);
  const rows = getData_(sh);
  const editable = new Set();
  rows.forEach(r=>{
    const col = String(r[map['columna']]).trim();
    const ok  = String(r[map['editable']]).trim().toLowerCase();
    if (col && (ok === 'si' || ok === 'sí' || ok === 'yes' || ok === 'true' || ok === 'x')){
      editable.add(col);
    }
  });
  // Asegura las mínimas por defecto
  DEFAULT_EDITABLE.forEach(c => editable.add(c));
  return editable;
}

/***** Protección de columnas en 📦 Cotizaciones_Master *****
 * Estrategia:
 * - Columnas MARCADAS COMO EDITABLES: se dejan SIN protección (todos pueden editar según sharing)
 * - Columnas NO EDITABLES: se crean protecciones por columna que solo permiten a ADMINs
 * Nota: Google Sheets no permite “lista de editores” dinámica sin recrear protecciones.
 ************************************************************/
function rebuildColumnProtections_(){
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(SHEET_MASTER) || getOrCreateSheet_(NICE_SHEET_MASTER);
  if (!sh) throw new Error('No se encontró la hoja Master');

  // Limpia protecciones previas creadas por este script
  const all = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
  all.forEach(p => {
    try {
      const desc = p.getDescription() || '';
      if (desc.startsWith('[cot-perm]')) p.remove();
    } catch(e){}
  });

  const headers = getHeaders_(sh);
  if (!headers || headers.length === 0) return;

  const editableSet = getEditableColumns_();
  const admins = getAdmins_();

  // Si no hay admins definidos, asegura al menos ADMIN_EMAIL
  const adminList = admins.length ? admins : (ADMIN_EMAIL ? [ADMIN_EMAIL] : []);

  headers.forEach((colName, idx) => {
    const colIndex = idx + 1;
    const isEditable = editableSet.has(String(colName));
    if (isEditable) {
      // Columna libre: sin protección
      return;
    }
    // Columna restringida: solo admins
    const range = sh.getRange(2, colIndex, Math.max(1, sh.getMaxRows()-1), 1); // desde fila 2
    const prot = range.protect();
    prot.setDescription(`[cot-perm] ${colName}`);
    prot.removeEditors(prot.getEditors()); // limpia primero
    if (adminList.length){
      prot.addEditors(adminList);
    }
    // Opcional: impedir que dominio la edite si es dominio compartido
    try { prot.setDomainEdit(false); } catch(e){}
  });

  SpreadsheetApp.getUi().alert('Permisos actualizados ✅');
}

/***** Menú/acciones de administración *****/
function menuRebuildPermissions_(){
  ensureUsersSheet_();
  rebuildColumnProtections_();
}

function adminAddUser(email, rol){
  ensureUsersSheet_();
  const sh = getOrCreateSheet_(SHEET_USERS) || getOrCreateSheet_(NICE_SHEET_USERS);
  const normalized = String(email||'').trim().toLowerCase();
  if (!normalized) throw new Error('Email requerido');
  const r = String(rol||'').trim().toLowerCase();
  if (!['admin','editor','viewer'].includes(r)) throw new Error('Rol inválido (admin|editor|viewer)');

  // dedupe simple
  const data = getData_(sh);
  const exists = data.some(row => String(row[0]||'').trim().toLowerCase() === normalized);
  if (!exists){
    sh.appendRow([normalized, r, '']);
  } else {
    // actualiza rol si ya existe
    const H = getHeaders_(sh);
    const map = idxMap_(H);
    const lr = sh.getLastRow();
    for (let i=2;i<=lr;i++){
      const v = String(sh.getRange(i,1).getValue()||'').trim().toLowerCase();
      if (v === normalized){
        sh.getRange(i,2).setValue(r);
        break;
      }
    }
  }
  SpreadsheetApp.getUi().alert(`Usuario ${normalized} → ${r} ✅`);
}

function adminRemoveUser(email){
  ensureUsersSheet_();
  const sh = getOrCreateSheet_(SHEET_USERS) || getOrCreateSheet_(NICE_SHEET_USERS);
  const normalized = String(email||'').trim().toLowerCase();
  const lr = sh.getLastRow();
  for (let i=lr;i>=2;i--){
    const v = String(sh.getRange(i,1).getValue()||'').trim().toLowerCase();
    if (v === normalized){
      sh.deleteRow(i);
    }
  }
  SpreadsheetApp.getUi().alert(`Usuario ${normalized} eliminado de la lista ✅`);
}
