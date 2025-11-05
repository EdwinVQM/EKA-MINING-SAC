/** Crear todas las hojas de catálogo si no existen (útil primera vez) */
function ensureCatalogSheets() {
  Object.keys(HEADERS).forEach(name => {
    const sh = SpreadsheetApp.getActive().getSheetByName(name) || SpreadsheetApp.getActive().insertSheet(name);
    if (sh.getLastRow() === 0) {
      const headers = HEADERS[name] || [];
      if (headers.length) sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  });
  return { ok:true };
}
