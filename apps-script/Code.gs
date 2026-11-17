/**
 * PROGRAMAS DE PRODUCCIÓN (Troqueles · Estructura · Muebles) · Backend v1
 * ---------------------------------------------------------------------------
 * Guarda los datos de la herramienta en este Google Sheet para que se vean igual en cualquier navegador
 * (Leonardo, operadores y la rutina del agente que sube el previo cada mañana).
 *
 * Hoja «Datos»: coleccion | clave | parte | json | actualizado (ms) | por
 *   colecciones: modelos · despiece · equiv · bcRutas · programas · registros · borradores
 *   Un documento grande se parte en varias filas (parte 0, 1, …) de hasta 45,000 caracteres.
 *   Borrar deja la fila con json vacío (marca) para que los otros navegadores también lo borren.
 *
 * Lecturas (GET, JSONP con «callback»):
 *   accion=ping                 → {ok, libro, filas}
 *   accion=todo                 → {ok, ahora, datos:{col:{clave:valor}}, borrados:{col:[claves]}}
 *   accion=cambios&desde=<ms>   → lo mismo, solo lo actualizado después de «desde»
 * Escrituras (POST, cuerpo JSON en texto):
 *   {accion:"guardar", cambios:[{col, clave, json:"<documento en JSON>"} | {col, clave, borrar:true}], por}
 *
 * Instalación: este script va dentro del Sheet (Extensiones → Apps Script).
 *   Implementar → Nueva implementación → Aplicación web · Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
 *   La URL /exec que da se pega en la herramienta (botón «Nube» del menú lateral).
 */

const HOJA = 'Datos';
const MAX = 45000;                 // caracteres por celda (el límite de Sheets es 50,000)
const COLS = ['modelos', 'despiece', 'equiv', 'bcRutas', 'programas', 'registros', 'borradores'];

function hoja_() {
  const l = SpreadsheetApp.getActiveSpreadsheet();
  let h = l.getSheetByName(HOJA);
  if (!h) {
    h = l.insertSheet(HOJA);
    h.getRange(1, 1, 1, 6).setValues([['coleccion', 'clave', 'parte', 'json', 'actualizado', 'por']]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  // clave y json siempre como texto: un pedazo de JSON que empiece con números o fechas no debe convertirse
  if (h.getRange('D2').getNumberFormat() !== '@') { h.getRange('B:B').setNumberFormat('@'); h.getRange('D:D').setNumberFormat('@'); }
  return h;
}

function responder_(p, out) {
  const txt = JSON.stringify(out);
  const cb = p.callback && /^[\w.$]+$/.test(p.callback) ? p.callback : null;
  return cb
    ? ContentService.createTextOutput(cb + '(' + txt + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    switch (p.accion) {
      case 'ping':    out = { ok: true, libro: SpreadsheetApp.getActiveSpreadsheet().getName(), filas: Math.max(0, hoja_().getLastRow() - 1), ahora: Date.now() }; break;
      case 'todo':    out = leer_(0, false); break;
      case 'cambios': out = leer_(Number(p.desde) || 0, true); break;
      default:        throw new Error('Acción desconocida: ' + p.accion);
    }
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return responder_(p, out);
}

function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.accion !== 'guardar') throw new Error('Acción desconocida: ' + b.accion);
    out = conCandado_(() => guardar_(b.cambios || [], String(b.por || '').slice(0, 60)));
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function conCandado_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(30000);
  try { return fn(); } finally { l.releaseLock(); }
}

/** Lee la hoja y arma los documentos. Con soloCambios, solo los actualizados después de «desde». */
function leer_(desde, soloCambios) {
  const h = hoja_(), n = h.getLastRow() - 1, ahora = Date.now();
  const datos = {}, borrados = {};
  COLS.forEach(c => { datos[c] = {}; borrados[c] = []; });
  if (n <= 0) return { ok: true, ahora, datos, borrados };
  const v = h.getRange(2, 1, n, 5).getValues();
  const docs = {};   // "col\u0001clave" → {col, clave, partes:[], t}
  for (const r of v) {
    const col = String(r[0]), clave = String(r[1]);
    if (!col || !clave) continue;
    const k = col + '\u0001' + clave, d = docs[k] || (docs[k] = { col, clave, partes: [], t: 0 });
    d.partes[Number(r[2]) || 0] = String(r[3] || '');
    d.t = Math.max(d.t, Number(r[4]) || 0);
  }
  for (const k in docs) {
    const d = docs[k];
    if (soloCambios && d.t <= desde) continue;
    if (!datos[d.col]) { datos[d.col] = {}; borrados[d.col] = []; }
    const json = d.partes.join('');
    if (!json) { borrados[d.col].push(d.clave); continue; }
    try { datos[d.col][d.clave] = JSON.parse(json); } catch (err) { /* documento a medio escribir: se ignora */ }
  }
  return { ok: true, ahora, datos, borrados };
}

/** Escribe documentos (partidos en filas) y marcas de borrado. Reusa las filas existentes de cada documento. */
function guardar_(cambios, por) {
  const h = hoja_(), n = h.getLastRow() - 1, ahora = Date.now();
  const filas = {};   // "col\u0001clave\u0001parte" → número de fila
  if (n > 0) h.getRange(2, 1, n, 3).getValues().forEach((r, i) => { filas[r[0] + '\u0001' + r[1] + '\u0001' + (Number(r[2]) || 0)] = i + 2; });
  const nuevas = [];
  let escritos = 0;
  for (const c of cambios) {
    if (!c || COLS.indexOf(c.col) < 0 || !c.clave) continue;
    const clave = String(c.clave), json = c.borrar ? '' : (typeof c.json === 'string' ? c.json : JSON.stringify(c.valor));
    const partes = [];
    for (let i = 0; i < Math.max(1, Math.ceil(json.length / MAX)); i++) partes.push(json.slice(i * MAX, (i + 1) * MAX));
    // las partes que sobran de una versión anterior más larga se vacían
    for (let i = partes.length; filas[c.col + '\u0001' + clave + '\u0001' + i]; i++) partes.push('');
    partes.forEach((txt, i) => {
      const fila = [c.col, clave, i, txt, ahora, por];
      const f = filas[c.col + '\u0001' + clave + '\u0001' + i];
      if (f) h.getRange(f, 1, 1, 6).setValues([fila]); else nuevas.push(fila);
    });
    escritos++;
  }
  if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, 6).setValues(nuevas);
  return { ok: true, escritos, ahora };
}
