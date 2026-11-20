/**
 * Mis Cuentas – recibe los pagos de Apple Pay desde el atajo y se los da a la app.
 * 1. Cambia CLAVE por una palabra secreta tuya (la misma que pondrás en el atajo y en la app).
 * 2. Implementar → Nueva implementación → Aplicación web
 *    Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
 */
const CLAVE = 'CAMBIA-ESTA-CLAVE';
const HOJA = 'Pagos';

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA);
  if (!h) {
    h = ss.insertSheet(HOJA);
    h.appendRow(['Recibido', 'Línea (fecha|importe|comercio)']);
    h.getRange('B:B').setNumberFormat('@');
  }
  return h;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// El atajo envía aquí cada pago (POST, formulario con t y linea)
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.t !== CLAVE) return json_({ ok: false, error: 'clave' });
  const linea = String(p.linea || '').trim();
  if (!linea) return json_({ ok: false, error: 'vacio' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    hoja_().appendRow([new Date(), "'" + linea]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// La app pide aquí los pagos nuevos (GET con t y desde)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.t !== CLAVE) return json_({ ok: false, error: 'clave' });
  const h = hoja_();
  const desde = Math.max(2, parseInt(p.desde || '2', 10) || 2);
  const ultima = h.getLastRow();
  const filas = [];
  if (ultima >= desde) {
    h.getRange(desde, 2, ultima - desde + 1, 1).getDisplayValues()
      .forEach((r, i) => filas.push({ n: desde + i, linea: String(r[0]) }));
  }
  return json_({ ok: true, filas: filas, siguiente: ultima + 1 });
}
