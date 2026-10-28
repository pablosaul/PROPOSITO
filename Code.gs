/**
 * Café Propósito — receptor de encuestas + API del dashboard
 * Pegar en: Google Sheet > Extensiones > Apps Script
 * Luego: Implementar > Nueva implementación > Aplicación web
 *   Ejecutar como: Yo   |   Quién tiene acceso: Cualquier persona
 */
const HOJA = 'Respuestas';
const TZ = 'America/Lima';
const DIAS = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
const HEAD = ['Fecha','Hora','Hora (0-23)','Día','Nombre','Contacto','Preferencia','Plato favorito',
              'Amabilidad','Rapidez','Ambiente','Producto nuevo','Qué mejorar','Origen'];

// Define tu clave del dashboard ejecutando UNA vez la función configurarClave()
function configurarClave() {
  PropertiesService.getScriptProperties().setProperty('DASH_KEY', 'CAMBIA-ESTA-CLAVE');
}

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#4b5330').setFontColor('#ffffff');
  }
  return sh;
}

// Evita que un texto del cliente se interprete como fórmula
function limpio_(v) {
  v = String(v == null ? '' : v).slice(0, 600);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function num_(v) { v = Number(v); return v >= 1 && v <= 5 ? v : ''; }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const now = new Date(); // la hora se toma del servidor (Lima), no del celular del cliente
    const dia = DIAS[Number(Utilities.formatDate(now, TZ, 'u')) % 7];
    hoja_().appendRow([
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(now, TZ, 'HH:mm:ss'),
      Number(Utilities.formatDate(now, TZ, 'H')),
      dia,
      limpio_(d.nombre), limpio_(d.contacto), limpio_(d.pref), limpio_(d.plato),
      num_(d.amabilidad), num_(d.rapidez), num_(d.ambiente),
      limpio_(d.nuevo), limpio_(d.mejora), limpio_(d.origen)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// El dashboard consulta: URL?key=TU_CLAVE  (no expone nombre ni contacto)
function doGet(e) {
  const clave = PropertiesService.getScriptProperties().getProperty('DASH_KEY');
  if (!clave || !e.parameter || e.parameter.key !== clave) return json_({ ok: false, error: 'unauthorized' });
  const sh = hoja_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: true, rows: [] });
  const v = sh.getRange(2, 1, n, HEAD.length).getDisplayValues();
  const rows = v.map(r => ({
    fecha: r[0], hora: r[1], h: Number(r[2]), dia: r[3], pref: r[6], plato: r[7],
    a: Number(r[8]), r: Number(r[9]), m: Number(r[10]), nuevo: r[11], mejora: r[12]
  }));
  return json_({ ok: true, rows });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
