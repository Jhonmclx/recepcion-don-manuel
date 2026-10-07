/**
 * Servidor de PEDIDOS A BODEGA — Mercados Don Manuel
 *
 * 1. Cree una Hoja de cálculo de Google nueva (sheets.new) y llámela "Pedidos Don Manuel".
 * 2. Extensiones > Apps Script. Borre lo que haya, pegue TODO este código y guarde (💾).
 * 3. Configuración del proyecto (engranaje ⚙) > Propiedades de la secuencia de comandos >
 *    Agregar propiedad:  CLAVE  =  (una clave que usted invente, ej: donmanuel2026)
 * 4. Implementar > Nueva implementación > tipo "Aplicación web":
 *       Ejecutar como: Yo        Quién tiene acceso: Cualquier persona
 *    Implementar, autorizar con su cuenta, y copiar la URL que termina en /exec
 */

var ENC = ['Id', 'Numero', 'Sede', 'Estado', 'Pidio', 'Creado', 'Nota', 'Alisto', 'Nota bodega', 'Despachado', 'Recibio', 'Recibido', 'Actualizado', 'Productos', 'Datos'];
var ENC_DET = ['Numero', 'Sede', 'Fecha pedido', 'Codigo', 'Producto', 'Pedido', 'Alistado', 'Faltante', 'Despachado'];

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function claveOk_(k) { var c = PropertiesService.getScriptProperties().getProperty('CLAVE'); return !!c && String(k) === String(c); }
function hoja_(n, enc) { var ss = SpreadsheetApp.getActive(); var s = ss.getSheetByName(n); if (!s) { s = ss.insertSheet(n); s.appendRow(enc); s.setFrozenRows(1); } return s; }
function stamp_() { var t = Date.now(); PropertiesService.getScriptProperties().setProperty('VERSION', String(t)); return t; }
function fecha_(ms) { return ms ? Utilities.formatDate(new Date(ms), 'America/Bogota', 'yyyy-MM-dd HH:mm') : ''; }

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!claveOk_(p.clave)) return out_({ ok: false, error: 'clave' });
  var ver = PropertiesService.getScriptProperties().getProperty('VERSION') || '0';
  if (p.accion === 'version' || p.accion === 'ping') return out_({ ok: true, version: ver });
  if (p.accion === 'pedidos') {
    var v = hoja_('Pedidos', ENC).getDataRange().getValues(); v.shift();
    var list = [];
    for (var i = v.length - 1; i >= 0 && list.length < 300; i--) { try { list.push(JSON.parse(v[i][14])); } catch (err) {} }
    return out_({ ok: true, version: ver, pedidos: list });
  }
  return out_({ ok: false, error: 'accion' });
}

function doPost(e) {
  var d; try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'datos' }); }
  if (!claveOk_(d.clave)) return out_({ ok: false, error: 'clave' });
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var s = hoja_('Pedidos', ENC);
    if (d.accion === 'crear') {
      var o = d.pedido; if (!o || !o.id) return out_({ ok: false, error: 'datos' });
      if (filaDe_(s, o.id) > 0) return out_({ ok: true, repetido: true, version: stamp_() });
      s.appendRow(fila_(o));
      return out_({ ok: true, version: stamp_() });
    }
    if (d.accion === 'actualizar') {
      var f = filaDe_(s, d.id); if (f < 0) return out_({ ok: false, error: 'no existe' });
      var cur = JSON.parse(s.getRange(f, 15).getValue());
      var patch = d.patch || {};
      for (var k in patch) cur[k] = patch[k];
      if (d.items) cur.items = d.items;               // lista completa de productos (alistados)
      cur.actualizado = Date.now();
      s.getRange(f, 1, 1, ENC.length).setValues([fila_(cur)]);
      if (patch.estado === 'despachado') detalle_(cur);
      return out_({ ok: true, version: stamp_(), pedido: cur });
    }
    if (d.accion === 'borrar') {
      var fb = filaDe_(s, d.id); if (fb > 0) s.deleteRow(fb);
      return out_({ ok: true, version: stamp_() });
    }
    return out_({ ok: false, error: 'accion' });
  } finally { lock.releaseLock(); }
}

function filaDe_(s, id) { var n = s.getLastRow(); if (n < 2) return -1; var ids = s.getRange(2, 1, n - 1, 1).getValues(); for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2; return -1; }
function fila_(o) {
  var resumen = (o.items || []).map(function (x) { return x.qty + ' ' + x.name + (x.alistado != null && x.alistado < x.qty ? ' (llegó ' + x.alistado + ')' : ''); }).join(' | ');
  return [o.id, o.numero, o.sede, o.estado, o.pidio || '', fecha_(o.creado), o.nota || '', o.alisto || '', o.notaBodega || '',
    fecha_(o.despachado), o.recibio || '', fecha_(o.recibido), fecha_(o.actualizado || o.creado), resumen, JSON.stringify(o)];
}
function detalle_(o) {
  var sd = hoja_('Detalle', ENC_DET);
  var rows = (o.items || []).map(function (x) { var al = x.alistado == null ? x.qty : x.alistado; return [o.numero, o.sede, fecha_(o.creado), x.code || '', x.name, x.qty, al, x.qty - al, fecha_(o.despachado)]; });
  if (rows.length) { var f = sd.getLastRow() + 1; sd.getRange(f, 4, rows.length, 1).setNumberFormat('@'); sd.getRange(f, 1, rows.length, rows[0].length).setValues(rows); }
}
