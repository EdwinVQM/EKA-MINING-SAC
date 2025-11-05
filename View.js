/** Funciones de lectura/consulta rápidas (placeholder) */
function listSheet(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return { ok:false, error:'Hoja no encontrada: '+name };
  const values = sh.getDataRange().getDisplayValues();
  return { ok:true, values };
}
