// Pajapp — backend en Google Apps Script (Extensiones > Apps Script del Sheet)
const SHEET = 'Pajereadas';
const VOTOS = 'Votos';
const USUARIOS = 'Usuarios';
const NOMBRES = ['Juancho', 'Puntano', 'Ivan', 'Mati', 'Alan', 'Bachi', 'Head'];
const HEAD_P = ['id', 'fecha', 'nombre', 'frase', 'cargadoPor', 'categoria', 'apelada', 'motivo'];
const HEAD_V = ['pajereadaId', 'votante', 'valor', 'tipo', 'fecha'];
const HEAD_U = ['nombre', 'pinHash', 'fecha'];
const COOLDOWN_MS = 10000; // 10 s entre pajereadas (para todos)

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
function usrSheet_() { return ensure_(USUARIOS, HEAD_U); }

// Correr a mano: crea/actualiza hojas y resumen
function setup() {
  pajSheet_();
  votSheet_();
  usrSheet_();
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
function err_(msg) { return json_({ ok: false, error: msg }); }

function rows_(sh, cols) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols).getValues();
}

function hash_(nombre, pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'pajapp|' + nombre + '|' + pin);
  return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function usuario_(nombre) {
  const u = rows_(usrSheet_(), 2);
  for (let i = 0; i < u.length; i++) if (u[i][0] === nombre) return { row: i + 2, hash: String(u[i][1]) };
  return null;
}

// Devuelve '' si el nombre+pin es valido, o el mensaje de error
function auth_(nombre, pin) {
  if (NOMBRES.indexOf(nombre) === -1) return 'nombre invalido';
  const u = usuario_(nombre);
  if (!u) return 'Primero identificate con tu nombre y PIN';
  if (u.hash !== hash_(nombre, String(pin || ''))) return 'PIN incorrecto';
  return '';
}

function state_() {
  const data = rows_(pajSheet_(), HEAD_P.length)
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: String(r[0]), fecha: new Date(r[1]).toISOString(), nombre: r[2], frase: r[3], por: r[4],
        categoria: r[5] || '', apelada: r[6] ? new Date(r[6]).toISOString() : '', motivo: String(r[7] || '')
      };
    });
  const votos = rows_(votSheet_(), HEAD_V.length)
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { id: String(r[0]), votante: r[1], valor: Number(r[2]), tipo: r[3] }; });
  const tomados = rows_(usrSheet_(), 1).map(function (r) { return r[0]; }).filter(String);
  return { ok: true, data: data, votos: votos, tomados: tomados, now: Date.now(), cooldown: COOLDOWN_MS };
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

// Lecturas por GET (sin datos sensibles en la URL)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action && p.action !== 'list') return err_('Version vieja de la pagina: recargala');
  try { return json_(state_()); } catch (err) { return err_(String(err)); }
}

// Escrituras por POST (body JSON en text/plain)
function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return err_('body invalido'); }
  try {
    const res = handle_(p);
    if (res) return err_(res);
    return json_(state_());
  } catch (err) {
    return err_(String(err));
  }
}

function handle_(p) {
  const action = p.action;

  if (action === 'claim') {
    if (NOMBRES.indexOf(p.nombre) === -1) return 'nombre invalido';
    if (!/^\d{4,8}$/.test(String(p.pin || ''))) return 'El PIN tiene que ser de 4 a 8 numeros';
    return withLock_(function () {
      const u = usuario_(p.nombre);
      if (!u) { usrSheet_().appendRow([p.nombre, hash_(p.nombre, p.pin), new Date()]); return ''; }
      if (u.hash !== hash_(p.nombre, p.pin)) return p.nombre + ' ya lo tiene otro. Si sos vos, pone tu PIN.';
      return '';
    });
  }

  if (action === 'add') {
    if (NOMBRES.indexOf(p.nombre) === -1) return 'nombre invalido';
    return withLock_(function () {
      const sh = pajSheet_();
      const last = sh.getLastRow();
      if (last >= 2) {
        const r = sh.getRange(last, 1, 1, HEAD_P.length).getValues()[0];
        const diff = Date.now() - new Date(r[1]).getTime();
        if (diff < COOLDOWN_MS) {
          const s = Math.ceil((COOLDOWN_MS - diff) / 1000);
          return 'COOLDOWN|' + s + '|' + r[2] + '|' + String(r[3] || '').slice(0, 80);
        }
      }
      sh.appendRow([Utilities.getUuid().slice(0, 8), new Date(), p.nombre,
        String(p.frase || '').slice(0, 300), String(p.por || '').slice(0, 40), String(p.categoria || '').slice(0, 40), '', '']);
      return '';
    });
  }

  // Todo lo demas requiere identidad
  const a = auth_(p.votante, p.pin);
  if (a) return a;

  if (action === 'delete') {
    return withLock_(function () {
      const sh = pajSheet_();
      const row = findRow_(sh, p.id);
      if (row < 0) return '';
      if (sh.getRange(row, 3).getValue() === p.votante) return 'No podes borrar tus propias pajereadas. Apela.';
      sh.deleteRow(row);
      const vs = votSheet_();
      const v = rows_(vs, 1);
      for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === String(p.id)) vs.deleteRow(i + 2);
      return '';
    });
  }

  if (action === 'vote') {
    const tipo = p.tipo === 'juicio' ? 'juicio' : 'voto';
    const valor = Math.max(-1, Math.min(1, Number(p.valor) || 0));
    return withLock_(function () {
      const sh = pajSheet_();
      const row = findRow_(sh, p.id);
      if (row < 0) return 'no existe';
      const r = sh.getRange(row, 1, 1, HEAD_P.length).getValues()[0];
      if (r[2] === p.votante) return 'No te podes votar a vos mismo';
      if (tipo === 'juicio' && !r[6]) return 'no esta en juicio';
      const vs = votSheet_();
      const v = rows_(vs, 4);
      for (let i = v.length - 1; i >= 0; i--) {
        if (String(v[i][0]) === String(p.id) && v[i][1] === p.votante && v[i][3] === tipo) vs.deleteRow(i + 2);
      }
      if (valor !== 0) vs.appendRow([p.id, p.votante, valor, tipo, new Date()]);
      return '';
    });
  }

  if (action === 'apelar') {
    return withLock_(function () {
      const sh = pajSheet_();
      const row = findRow_(sh, p.id);
      if (row < 0) return 'no existe';
      const r = sh.getRange(row, 1, 1, HEAD_P.length).getValues()[0];
      if (r[2] !== p.votante) return 'Solo el acusado puede apelar';
      if (r[6]) return 'Ya fue apelada';
      const motivo = String(p.motivo || '').trim().slice(0, 500);
      if (motivo.length < 5) return 'Explica por que apelas (minimo 5 letras)';
      sh.getRange(row, 7, 1, 2).setValues([[new Date(), motivo]]);
      return '';
    });
  }

  return 'accion desconocida';
}
