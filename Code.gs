// Pajapp — backend en Google Apps Script (Extensiones > Apps Script del Sheet)
const SHEET = 'Pajereadas';
const VOTOS = 'Votos';
const NOMBRES = ['Juancho', 'Puntano', 'Ivan', 'Mati', 'Alan', 'Bachi', 'Head'];
const HEAD_P = ['id', 'fecha', 'nombre', 'frase', 'cargadoPor', 'categoria', 'apelada'];
const HEAD_V = ['pajereadaId', 'votante', 'valor', 'tipo', 'fecha'];

function ensure_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const cur = sh.getLastRow() ? sh.getRange(1, 1, 1, headers.length).getValues()[0] : [];
  if (cur.join('|') !== headers.join('|')) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function pajSheet_() { return ensure_(SHEET, HEAD_P); }
function votSheet_() { return ensure_(VOTOS, HEAD_V); }

// Correr a mano: crea/actualiza hojas y resumen
function setup() {
  pajSheet_();
  votSheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let res = ss.getSheetByName('Resumen');
  if (!res) res = ss.insertSheet('Resumen');
  res.clear();
  res.appendRow(['Nombre', 'Total cargadas', 'Hoy']);
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

function rows_(sh, cols) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols).getValues();
}

function state_() {
  const data = rows_(pajSheet_(), HEAD_P.length)
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: String(r[0]), fecha: new Date(r[1]).toISOString(), nombre: r[2], frase: r[3], por: r[4],
        categoria: r[5] || '', apelada: r[6] ? new Date(r[6]).toISOString() : ''
      };
    });
  const votos = rows_(votSheet_(), HEAD_V.length)
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { id: String(r[0]), votante: r[1], valor: Number(r[2]), tipo: r[3] }; });
  return { ok: true, data: data, votos: votos };
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function err_(msg) { return json_({ ok: false, error: msg }); }

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'list';
  try {
    if (action === 'add') {
      if (NOMBRES.indexOf(p.nombre) === -1) return err_('nombre invalido');
      withLock_(function () {
        pajSheet_().appendRow([Utilities.getUuid().slice(0, 8), new Date(), p.nombre,
          String(p.frase || '').slice(0, 300), String(p.por || '').slice(0, 40), String(p.categoria || '').slice(0, 40), '']);
      });
    } else if (action === 'delete') {
      withLock_(function () {
        const sh = pajSheet_();
        const row = findRow_(sh, p.id);
        if (row > 0) sh.deleteRow(row);
        const vs = votSheet_();
        const v = rows_(vs, 1);
        for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === String(p.id)) vs.deleteRow(i + 2);
      });
    } else if (action === 'vote') {
      const tipo = p.tipo === 'juicio' ? 'juicio' : 'voto';
      const valor = Math.max(-1, Math.min(1, Number(p.valor) || 0));
      if (NOMBRES.indexOf(p.votante) === -1) return err_('votante invalido');
      const res = withLock_(function () {
        const sh = pajSheet_();
        const row = findRow_(sh, p.id);
        if (row < 0) return 'no existe';
        const r = sh.getRange(row, 1, 1, HEAD_P.length).getValues()[0];
        if (r[2] === p.votante) return 'no te podes votar a vos mismo';
        if (tipo === 'juicio' && !r[6]) return 'no esta en juicio';
        const vs = votSheet_();
        const v = rows_(vs, 4);
        for (let i = v.length - 1; i >= 0; i--) {
          if (String(v[i][0]) === String(p.id) && v[i][1] === p.votante && v[i][3] === tipo) vs.deleteRow(i + 2);
        }
        if (valor !== 0) vs.appendRow([p.id, p.votante, valor, tipo, new Date()]);
        return '';
      });
      if (res) return err_(res);
    } else if (action === 'apelar') {
      const res = withLock_(function () {
        const sh = pajSheet_();
        const row = findRow_(sh, p.id);
        if (row < 0) return 'no existe';
        const r = sh.getRange(row, 1, 1, HEAD_P.length).getValues()[0];
        if (r[2] !== p.votante) return 'solo el acusado puede apelar';
        if (r[6]) return 'ya fue apelada';
        sh.getRange(row, 7).setValue(new Date());
        return '';
      });
      if (res) return err_(res);
    }
    return json_(state_());
  } catch (err) {
    return err_(String(err));
  }
}
