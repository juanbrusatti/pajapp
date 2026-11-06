// Pajapp — backend en Google Apps Script (pegar en Extensiones > Apps Script del Sheet)
const SHEET = 'Pajereadas';
const NOMBRES = ['Juancho', 'Puntano', 'Ivan', 'Mati', 'Alan'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(['id', 'fecha', 'nombre', 'frase', 'cargadoPor']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
  }
  return sh;
}

// Correr una vez a mano: crea hojas y resumen
function setup() {
  getSheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let res = ss.getSheetByName('Resumen');
  if (!res) res = ss.insertSheet('Resumen');
  res.clear();
  res.appendRow(['Nombre', 'Total', 'Hoy']);
  NOMBRES.forEach(function (n, i) {
    const r = i + 2;
    res.getRange(r, 1).setValue(n);
    res.getRange(r, 2).setFormula('=COUNTIF(Pajereadas!C:C,A' + r + ')');
    res.getRange(r, 3).setFormula('=COUNTIFS(Pajereadas!C:C,A' + r + ',Pajereadas!B:B,">="&TODAY())');
  });
  res.getRange('A1:C1').setFontWeight('bold');
  const def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function list_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { id: String(r[0]), fecha: new Date(r[1]).toISOString(), nombre: r[2], frase: r[3], por: r[4] };
    });
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'list';
  try {
    if (action === 'add') {
      if (NOMBRES.indexOf(p.nombre) === -1) return json_({ ok: false, error: 'nombre invalido' });
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const id = Utilities.getUuid().slice(0, 8);
        getSheet_().appendRow([id, new Date(), p.nombre, String(p.frase || '').slice(0, 300), String(p.por || '').slice(0, 40)]);
      } finally { lock.releaseLock(); }
    } else if (action === 'delete') {
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const sh = getSheet_();
        const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
        for (let i = ids.length - 1; i >= 1; i--) {
          if (String(ids[i][0]) === String(p.id)) { sh.deleteRow(i + 1); break; }
        }
      } finally { lock.releaseLock(); }
    }
    return json_({ ok: true, data: list_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
