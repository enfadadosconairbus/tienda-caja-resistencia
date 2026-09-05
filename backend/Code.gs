/* =============================================================================
   Code.gs — Backend Apps Script V4.2 · Tienda Caja de Resistencia (Airbus 2026)
   -----------------------------------------------------------------------------
   Panel de operaciones sobre Google Sheet, con estética "centro de operaciones":
     · DASHBOARD con KPIs en tarjetas de color + 2 gráficos (dónut y barras).
     · Conciliación bancaria automática (código AIR26-XXXXX + importe exacto).
     · LOTES y PROVEEDOR por PRODUCTO + SKU + TALLA.
     · Estados con formato condicional por color · validaciones · hoja LOG.
     · Emails (recibido / confirmado / listo) · Exportar a Excel (.xlsx).

   Puesta en marcha:
     1) Pega este fichero en Apps Script del Sheet (cuenta operativa).
     2) Menú 👕 Tienda Airbus 2026 → 🛠️ Preparar backend (setup).
     3) Revisa CONFIG: IBAN y BENEFICIARIO reales.
     4) Implementar → Aplicación web (Ejecutar como: tú · Acceso: cualquiera)
        → copia la URL /exec (al Worker).
     5) Menú → 🔑 Mostrar TOKEN backend (al Worker, nunca a GitHub).

   El navegador NUNCA decide el precio: se recalcula aquí contra CATALOGO.
   ========================================================================== */

var SH = {
  HOWTO: '00_HOW_TO',
  DASH: '01_DASHBOARD',
  CONFIG: 'CONFIG',
  CATALOGO: 'CATALOGO',
  PEDIDOS: 'PEDIDOS',
  LINEAS: 'LINEAS_PEDIDO',
  BANCO: 'MOVIMIENTOS_BANCO',
  LOTES: 'LOTES',
  PROVEEDOR: 'PROVEEDOR',
  LOG: 'LOG',
  DATA: '_PANEL_DATA'   // hoja oculta: datos de los gráficos
};

var HEAD = {
  PEDIDOS: ['ID', 'FECHA_PEDIDO', 'NOMBRE', 'APELLIDOS', 'EMAIL', 'TELEFONO',
            'UNIDADES', 'PRODUCTOS_EUR', 'APORTACION_EUR', 'TOTAL_EUR', 'ESTADO',
            'CLIENT_REQUEST_ID', 'RECOGIDA', 'CADUCA', 'FECHA_CONFIRMADO',
            'FECHA_LISTO', 'FECHA_ENTREGADO', 'SITE', 'TARDIO'],
  LINEAS: ['ID', 'FECHA_PEDIDO', 'PRODUCTO', 'SKU', 'TALLA', 'CANTIDAD', 'LOTE'],
  CATALOGO: ['ACTIVO', 'PRODUCTO', 'SKU', 'TALLA', 'MEDIDAS', 'PRECIO', 'COSTE', 'APORTE_CAJA'],
  BANCO: ['FECHA', 'CONCEPTO', 'IMPORTE', 'REFERENCIA', 'PEDIDO_DETECTADO', 'RESULTADO', 'PROCESADO', 'FECHA_CONCILIACION'],
  LOTES: ['LOTE', 'FECHA_GENERACION', 'PRODUCTO', 'SKU', 'TALLA', 'CANTIDAD', 'ESTADO', 'FECHA_RECEPCION'],
  PROVEEDOR: ['LOTE', 'PRODUCTO', 'SKU', 'TALLA', 'CANTIDAD'],
  LOG: ['TIMESTAMP', 'TIPO', 'DETALLE']
};

// Flujo de estados de un pedido.
var ESTADOS = ['PENDIENTE_PAGO', 'PAGO_CONCILIADO', 'ENVIADO_PROVEEDOR', 'RECIBIDO', 'LISTO_RECOGIDA', 'ENTREGADO', 'CADUCADO'];
var ESTADOS_PAGADOS = ['PAGO_CONCILIADO', 'ENVIADO_PROVEEDOR', 'RECIBIDO', 'LISTO_RECOGIDA', 'ENTREGADO'];

// Marca "tardío" (columna TARDIO de PEDIDOS). Es INDEPENDIENTE del ESTADO: un pedido
// tardío se cobra y concilia con normalidad, pero NO entra solo al proveedor. Valores:
//   RETENIDO = pedido tras la fecha límite; se retiene de producción hasta hablar con el proveedor.
//   LIBERADO = ya hablado con el proveedor y confirmado que llega a tiempo → sí produce.
//   ''       = pedido a tiempo, flujo normal.
// Quien decide RETENIDO es el BACKEND al crear el pedido (CONFIG → AVISO_FECHA_LIMITE),
// no el navegador. Solo se retiene en 📦 Generar pedido a proveedor (ver más abajo).
var TARDIO_RETENIDO = 'RETENIDO';
var TARDIO_LIBERADO = 'LIBERADO';

// Paleta (siguiendo la estela del Excel de referencia, con el naranja de marca).
var COL = {
  ink: '#211E1A', orange: '#C75B12', orangeDark: '#9F4309', slate: '#596573',
  paper: '#FFFDF9', cream: '#FBF7EF', line: '#E7E0D5', white: '#FFFFFF',
  tGray: '#F5F6F7', tGreen: '#DFF1E8', tGreen2: '#E2F1E9', tBlue: '#E2EEF7',
  tBlue2: '#E7EEF6', tPeach: '#F6E1D2', tRed: '#F7E1DE', tYellow: '#FAEBC9'
};
var FMT_FECHA = 'dd/MM/yyyy HH:mm';
var FMT_EUR = '#,##0.00 €';

/* ===========================  ENDPOINT WEB  ================================ */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (!tokenValido(data.token)) return jsonOut({ ok: false, error: 'No autorizado' });
    if (data.action === 'crear_pedido') return jsonOut(crearPedido(data));
    if (data.action === 'estado') return jsonOut(estadoPublico());
    return jsonOut({ ok: false, error: 'Acción no reconocida' });
  } catch (err) {
    return jsonOut({ ok: false, error: 'Error del servidor: ' + err });
  }
}
function doGet() { return jsonOut({ ok: true, service: 'tienda-airbus', version: 'V4.2' }); }
function jsonOut(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ===========================  CREAR PEDIDO  =============================== */

function crearPedido(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActive();
    var cfg = leerConfig();

    var crid = String(data.client_request_id || '');
    if (crid) { var ex = buscarPorCRID(ss, crid); if (ex) return respuestaPedido(ex.id, ex.total, cfg); }

    var catalogo = leerCatalogo(ss);
    var lineas = [], unidades = 0, productos = 0;
    (data.lineas || []).forEach(function (l) {
      var item = catalogo[String(l.sku)];
      if (!item || item.activo !== true) return;
      var cant = Math.max(0, Math.floor(Number(l.cantidad) || 0));
      if (cant <= 0) return;
      unidades += cant; productos += item.precio * cant;
      lineas.push({ producto: item.producto, sku: item.sku, talla: item.talla, cantidad: cant });
    });
    var aportacion = Math.max(0, Number(data.aportacion) || 0);

    // Se permite aportar SIN camiseta (donación pura): basta con que haya aportación.
    // Solo se rechaza si no hay ni camisetas válidas ni aportación.
    if (!lineas.length && aportacion <= 0) return { ok: false, error: 'Indica una aportación o añade una camiseta.' };

    var maxUds = Number(cfg.MAX_UNIDADES || 20);
    if (unidades > maxUds) return { ok: false, error: 'Máximo ' + maxUds + ' unidades por pedido.' };

    var total = productos + aportacion;

    var c = (data.cliente || {});
    var nombre = limpiar(c.nombre, 60), apellidos = limpiar(c.apellidos, 100);
    var email = limpiar(c.email, 120), telefono = limpiar(c.telefono, 30);
    var site = limpiar(data.site, 40);
    if (!nombre || !apellidos || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Datos de cliente incompletos.' };

    var id = siguienteId(cfg);
    var ahora = new Date();
    var caduca = new Date(ahora.getTime() + Number(cfg.CADUCIDAD_HORAS || 12) * 3600 * 1000);
    var recogida = limpiar(data.recogida || cfg.RECOGIDA || '', 200);

    // ¿Pedido tardío? Lo decide el servidor con SU fecha límite (CONFIG), nunca el
    // navegador. Tardío = RETENIDO: se cobra y concilia igual, pero se retiene de
    // producción hasta que confirmemos plazos con el proveedor.
    var tardio = esTardio(cfg, ahora) ? TARDIO_RETENIDO : '';

    ss.getSheetByName(SH.PEDIDOS).appendRow([id, ahora, nombre, apellidos, email, telefono,
      unidades, productos, aportacion, total, 'PENDIENTE_PAGO', crid, recogida, caduca, '', '', '', site, tardio]);
    var hojaLineas = ss.getSheetByName(SH.LINEAS);
    lineas.forEach(function (l) { hojaLineas.appendRow([id, ahora, l.producto, l.sku, l.talla, l.cantidad, '']); });

    try { emailPedidoRecibido(email, id, nombre, lineas, productos, aportacion, total, cfg); }
    catch (e) { registrarLog(ss, 'EMAIL_ERROR', 'recibido ' + id + ': ' + e); }
    registrarLog(ss, 'PEDIDO', id + ' · ' + unidades + ' uds · ' + eur(total) + (tardio ? ' · TARDÍO (retenido de producción)' : ''));
    return respuestaPedido(id, total, cfg);
  } finally { lock.releaseLock(); }
}

// ¿Se creó el pedido a partir de la fecha límite? Fuente: CONFIG → AVISO_FECHA_LIMITE
// (misma fecha que la del aviso de la web, pero aquí manda el servidor). Admite el
// valor como texto ISO ('2026-09-06T21:00:00+02:00') o como Date si Sheets lo parseó.
// Sin fecha válida en CONFIG, ningún pedido se marca tardío (comportamiento seguro).
function esTardio(cfg, ahora) {
  var raw = cfg && cfg.AVISO_FECHA_LIMITE;
  if (!raw) return false;
  var limite = (raw instanceof Date) ? raw : new Date(String(raw));
  if (isNaN(limite.getTime())) return false;
  return (ahora || new Date()) >= limite;
}
function respuestaPedido(id, total, cfg) {
  return { ok: true, order_id: id, total: Number(total), beneficiario: cfg.BENEFICIARIO || '', iban: cfg.IBAN || '', concepto: id };
}

// Camisetas pedidas (unidades, excluyendo caducados) para el aviso de stock de la web.
function estadoPublico() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.PEDIDOS), H = HEAD.PEDIDOS;
  var last = sh.getLastRow(), total = 0;
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, H.length).getValues();
    var iU = H.indexOf('UNIDADES'), iE = H.indexOf('ESTADO');
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][iE]) === 'CADUCADO') continue;
      total += Number(vals[r][iU]) || 0;
    }
  }
  return { ok: true, camisetas: total };
}

/* ===========================  CONCILIACIÓN BANCARIA  ===================== */

function conciliarBanco() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SH.BANCO);
  var last = sh.getLastRow();
  if (last < 2) { ui().alert('No hay movimientos en ' + SH.BANCO + '. Pega el extracto (FECHA, CONCEPTO, IMPORTE, REFERENCIA).'); return; }

  var H = HEAD.BANCO, col = function (k) { return H.indexOf(k); };
  var rango = sh.getRange(2, 1, last - 1, H.length), vals = rango.getValues();
  var pedidos = indicePedidos(ss), cfg = leerConfig();
  var conc = 0, rev = 0, ya = 0, ahora = new Date();

  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][col('PROCESADO')]).toUpperCase() === 'SI') continue;
    var concepto = String(vals[r][col('CONCEPTO')] || '');
    var importe = parseImporte(vals[r][col('IMPORTE')]);
    var id = detectarId(concepto, cfg);
    if (!id) { vals[r][col('RESULTADO')] = 'REVISAR_SIN_CODIGO'; rev++; continue; }
    var p = pedidos[id];
    if (!p) { vals[r][col('RESULTADO')] = 'REVISAR_CODIGO_INEXISTENTE'; vals[r][col('PEDIDO_DETECTADO')] = id; rev++; continue; }
    vals[r][col('PEDIDO_DETECTADO')] = id;
    if (ESTADOS_PAGADOS.indexOf(p.estado) >= 0) { vals[r][col('RESULTADO')] = 'YA_PAGADO'; vals[r][col('PROCESADO')] = 'SI'; vals[r][col('FECHA_CONCILIACION')] = ahora; ya++; continue; }
    if (p.estado === 'CADUCADO') { vals[r][col('RESULTADO')] = 'REVISAR_CADUCADO'; rev++; continue; }
    if (Math.round(importe * 100) === Math.round(p.total * 100)) {
      marcarPedidoPagado(ss, p, cfg);
      vals[r][col('RESULTADO')] = 'PAGO_CONCILIADO'; vals[r][col('PROCESADO')] = 'SI'; vals[r][col('FECHA_CONCILIACION')] = ahora;
      pedidos[id].estado = 'PAGO_CONCILIADO'; conc++;
    } else { vals[r][col('RESULTADO')] = 'REVISAR_IMPORTE (' + eur(importe) + ' vs ' + eur(p.total) + ')'; rev++; }
  }
  rango.setValues(vals);
  registrarLog(ss, 'CONCILIACION', 'conciliados ' + conc + ' · revisar ' + rev + ' · ya ' + ya);
  refrescarDashboard();
  ui().alert('Conciliación terminada.\n\nConciliados: ' + conc + '\nYa pagados: ' + ya + '\nPara revisar: ' + rev);
}

function marcarPedidoPagado(ss, p, cfg) {
  var sh = ss.getSheetByName(SH.PEDIDOS), H = HEAD.PEDIDOS;
  sh.getRange(p.fila, H.indexOf('ESTADO') + 1).setValue('PAGO_CONCILIADO');
  sh.getRange(p.fila, H.indexOf('FECHA_CONFIRMADO') + 1).setValue(new Date());
  try { emailPagoConfirmado(p.email, p.id, p.nombre, lineasDePedido(ss, p.id), p.productos, p.aportacion, p.total, cfg, p.site); return true; }
  catch (e) { registrarLog(ss, 'EMAIL_ERROR', 'confirmado ' + p.id + ': ' + e); return false; }
}

/* ===========================  LOTES / PROVEEDOR  ======================== */

function generarPedidoProveedor() {
  var ss = SpreadsheetApp.getActive(), pedidos = indicePedidos(ss);

  // UN LOTE POR SITE: agrupamos los pedidos PAGO_CONCILIADO por su site, para que
  // el proveedor pueda producir y separar en cajas distintas por site.
  var ORDEN = ['albacete','cádiz','cadiz','getafe','illescas','san pablo','tablada'];
  var porSite = {}, retenidos = 0;
  for (var id in pedidos) {
    if (pedidos[id].estado !== 'PAGO_CONCILIADO') continue;
    // Pedidos tardíos RETENIDOS: se saltan (no entran solos a producción). Se
    // incluyen cuando se liberan a mano (🕒 Liberar pedidos tardíos), tras hablar
    // con el proveedor. LIBERADO y '' sí entran.
    if (pedidos[id].tardio === TARDIO_RETENIDO) { retenidos++; continue; }
    var nombre = String(pedidos[id].site || '').trim() || 'Sin site';
    var k = nombre.toLowerCase();
    if (!porSite[k]) porSite[k] = { nombre: nombre, ids: [] };
    porSite[k].ids.push(id);
  }
  var notaReten = retenidos ? ('\n\n🕒 ' + retenidos + ' pedido(s) TARDÍOS retenidos, NO incluidos en producción. ' +
    'Confirma plazos con el proveedor y libéralos con "🕒 Liberar pedidos tardíos (selección)".') : '';
  var claves = Object.keys(porSite).sort(function (a, b) {
    var ia = ORDEN.indexOf(a), ib = ORDEN.indexOf(b);
    if (ia < 0) ia = 99; if (ib < 0) ib = 99;
    return (ia - ib) || a.localeCompare(b);
  });
  if (!claves.length) { ui().alert('No hay pedidos listos para producir.' + notaReten); return; }

  var shL = ss.getSheetByName(SH.LINEAS), H = HEAD.LINEAS, lastL = shL.getLastRow();
  var lin = (lastL >= 2) ? shL.getRange(2, 1, lastL - 1, H.length).getValues() : [];
  var shLot = ss.getSheetByName(SH.LOTES), shProv = ss.getSheetByName(SH.PROVEEDOR);
  var shP = ss.getSheetByName(SH.PEDIDOS), HP = HEAD.PEDIDOS, ahora = new Date();

  var resumen = [], nLotes = 0, udsTotal = 0;
  claves.forEach(function (ck) {
    var grupo = porSite[ck], ids = grupo.ids, agg = {}, filas = [];
    for (var i = 0; i < lin.length; i++) {
      var row = lin[i], pid = String(row[H.indexOf('ID')]);
      if (ids.indexOf(pid) < 0) continue;
      if (String(row[H.indexOf('LOTE')])) continue;   // línea ya loteada: se salta
      var key = row[H.indexOf('PRODUCTO')] + '||' + row[H.indexOf('SKU')] + '||' + row[H.indexOf('TALLA')];
      if (!agg[key]) agg[key] = { producto: row[H.indexOf('PRODUCTO')], sku: row[H.indexOf('SKU')], talla: row[H.indexOf('TALLA')], cantidad: 0 };
      agg[key].cantidad += Number(row[H.indexOf('CANTIDAD')]) || 0;
      filas.push(i + 2);
    }
    var keys = Object.keys(agg);
    if (!keys.length) return;   // este site ya estaba loteado

    var lote = nuevoLoteId(ss) + '-' + slugSite(grupo.nombre);
    keys.forEach(function (kk) { var a = agg[kk];
      shLot.appendRow([lote, ahora, a.producto, a.sku, a.talla, a.cantidad, 'PEDIDO', '']);
      shProv.appendRow([lote, a.producto, a.sku, a.talla, a.cantidad]);
    });
    filas.forEach(function (f) { shL.getRange(f, H.indexOf('LOTE') + 1).setValue(lote); });
    ids.forEach(function (pid2) { shP.getRange(pedidos[pid2].fila, HP.indexOf('ESTADO') + 1).setValue('ENVIADO_PROVEEDOR'); });

    var uds = keys.reduce(function (s, kk) { return s + agg[kk].cantidad; }, 0);
    nLotes++; udsTotal += uds;
    resumen.push('• ' + grupo.nombre + ' → ' + lote + '  (' + uds + ' uds · ' + keys.length + ' líneas)');
    registrarLog(ss, 'LOTE', lote + ' · ' + grupo.nombre + ' · ' + keys.length + ' líneas · ' + uds + ' uds');
  });

  if (!nLotes) { ui().alert('Los pedidos pagados ya estaban loteados.' + notaReten); return; }
  regenerarResumenProveedor(ss);
  refrescarDashboard();
  ui().alert('Generados ' + nLotes + ' lote(s) — uno por site — · ' + udsTotal + ' uds en total:\n\n' + resumen.join('\n') +
    '\n\nDetalle en la hoja PROVEEDOR (ordénala por LOTE) y totales por talla × site en RESUMEN_PROVEEDOR. Las dos se exportan con 📗 Exportar a Excel.' + notaReten);
}

// Ítem de menú: reconstruye RESUMEN_PROVEEDOR sin generar lotes nuevos (útil tras
// editar/renombrar lotes a mano, p. ej. añadir el site a un lote antiguo).
function refrescarResumenProveedor() {
  regenerarResumenProveedor(SpreadsheetApp.getActive());
  ui().alert('RESUMEN_PROVEEDOR actualizado (totales por talla × site).');
}

// Pivote para el proveedor: filas = TALLA, columnas = cada SITE (del sufijo del LOTE),
// con TOTAL por talla (última columna) y TOTAL por site (última fila) + total general.
// Se regenera entero desde la hoja PROVEEDOR, así que siempre refleja todos los lotes.
function regenerarResumenProveedor(ss) {
  var shProv = ss.getSheetByName(SH.PROVEEDOR); if (!shProv) return;
  var HP = HEAD.PROVEEDOR, last = shProv.getLastRow();
  var rows = (last >= 2) ? shProv.getRange(2, 1, last - 1, HP.length).getValues() : [];

  var ORDEN_TALLAS = ['XS','S','M','L','XL','2XL','3XL','4XL','5XL'];
  var sites = [], tallas = [], mat = {}, totSite = {}, totTalla = {}, gran = 0;
  rows.forEach(function (r) {
    var lote = String(r[HP.indexOf('LOTE')] || '');
    if (!lote) return;
    var site = lote.split('-').pop();               // sufijo del lote = site
    var talla = String(r[HP.indexOf('TALLA')] || ''); if (!talla) return;
    var cant = Number(r[HP.indexOf('CANTIDAD')]) || 0;
    if (sites.indexOf(site) < 0) sites.push(site);
    if (tallas.indexOf(talla) < 0) tallas.push(talla);
    mat[talla] = mat[talla] || {};
    mat[talla][site] = (mat[talla][site] || 0) + cant;
    totSite[site] = (totSite[site] || 0) + cant;
    totTalla[talla] = (totTalla[talla] || 0) + cant;
    gran += cant;
  });

  sites.sort();
  tallas.sort(function (a, b) {
    var ia = ORDEN_TALLAS.indexOf(a), ib = ORDEN_TALLAS.indexOf(b);
    if (ia < 0) ia = 99; if (ib < 0) ib = 99;
    return (ia - ib) || a.localeCompare(b);
  });

  var sh = ss.getSheetByName('RESUMEN_PROVEEDOR') || ss.insertSheet('RESUMEN_PROVEEDOR');
  sh.clearContents();
  if (!tallas.length) { sh.getRange(1, 1).setValue('Sin datos de proveedor todavía.'); return; }

  var header = ['TALLA'].concat(sites).concat(['TOTAL']);
  var out = [header];
  tallas.forEach(function (t) {
    var fila = [t];
    sites.forEach(function (s) { fila.push((mat[t] && mat[t][s]) || 0); });
    fila.push(totTalla[t] || 0);
    out.push(fila);
  });
  var filaTotal = ['TOTAL'];
  sites.forEach(function (s) { filaTotal.push(totSite[s] || 0); });
  filaTotal.push(gran);
  out.push(filaTotal);

  sh.getRange(1, 1, out.length, header.length).setValues(out);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');                 // cabecera
  sh.getRange(out.length, 1, 1, header.length).setFontWeight('bold');        // fila TOTAL
  sh.getRange(1, header.length, out.length, 1).setFontWeight('bold');        // columna TOTAL
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, header.length); } catch (e) {}

  // --- Direcciones de envío por site (fuente única: CONFIG, clave ENVIO_<SITE>) ---
  // El proveedor hace el envío, así que replicamos aquí la dirección de cada site
  // presente en los datos. Si falta en CONFIG, la sembramos como placeholder para
  // que solo haya que rellenar el valor y volver a refrescar.
  var cfgSh = ss.getSheetByName(SH.CONFIG);
  if (cfgSh) {
    var existentes = {};
    if (cfgSh.getLastRow() >= 2)
      cfgSh.getRange(2, 1, cfgSh.getLastRow() - 1, 1).getValues()
        .forEach(function (r) { if (r[0]) existentes[String(r[0]).trim()] = true; });
    var faltan = [];
    sites.forEach(function (s) {
      var k = envioKeySite(s);
      if (!existentes[k]) { faltan.push([k, '[COMPLETAR dirección de envío · ' + siteNombre(s) + ']']); existentes[k] = true; }
    });
    if (faltan.length) cfgSh.getRange(cfgSh.getLastRow() + 1, 1, faltan.length, 2).setValues(faltan);
  }
  var cfg = leerConfig();
  var base = out.length + 2;   // dos filas debajo del pivote
  sh.getRange(base, 1).setValue('DIRECCIONES DE ENVÍO — el proveedor envía a cada site (fuente: CONFIG)').setFontWeight('bold');
  var dir = sites.map(function (s) { return [siteNombre(s), String(cfg[envioKeySite(s)] || '')]; });
  if (dir.length) {
    sh.getRange(base + 1, 1, dir.length, 2).setValues(dir);
    sh.getRange(base + 1, 1, dir.length, 1).setFontWeight('bold');   // nombre del site (la dirección desborda hacia la derecha)
  }
}

// Reconstruye la hoja PROVEEDOR desde la ÚNICA fuente de verdad: las líneas de
// LINEAS_PEDIDO que ya tienen LOTE. Úsalo tras editar tallas o cantidades a mano
// (PROVEEDOR solo se "iba añadiendo", por eso las ediciones manuales la descuadran).
// Deja PROVEEDOR = suma exacta de LINEAS-con-lote, sincroniza CANTIDAD en LOTES sin
// perder su ESTADO/FECHA_RECEPCION, y refresca RESUMEN_PROVEEDOR.
function reconstruirProveedorDesdeLineas() {
  var ss = SpreadsheetApp.getActive();
  var shL = ss.getSheetByName(SH.LINEAS), H = HEAD.LINEAS, lastL = shL.getLastRow();
  var lin = (lastL >= 2) ? shL.getRange(2, 1, lastL - 1, H.length).getValues() : [];

  // Agregar por LOTE|PRODUCTO|SKU|TALLA (solo líneas con LOTE).
  var agg = {}, orden = [];
  lin.forEach(function (r) {
    var lote = String(r[H.indexOf('LOTE')] || '').trim(); if (!lote) return;
    var prod = r[H.indexOf('PRODUCTO')], sku = r[H.indexOf('SKU')], talla = r[H.indexOf('TALLA')];
    var k = lote + '||' + prod + '||' + sku + '||' + talla;
    if (!(k in agg)) { agg[k] = { lote: lote, producto: prod, sku: sku, talla: talla, cantidad: 0 }; orden.push(k); }
    agg[k].cantidad += Number(r[H.indexOf('CANTIDAD')]) || 0;
  });

  // Reescribir PROVEEDOR entera.
  var shProv = ss.getSheetByName(SH.PROVEEDOR), HP = HEAD.PROVEEDOR, lastP = shProv.getLastRow();
  var out = orden.map(function (k) { var a = agg[k]; return [a.lote, a.producto, a.sku, a.talla, a.cantidad]; });
  if (lastP >= 2) shProv.getRange(2, 1, lastP - 1, HP.length).clearContent();
  if (out.length) shProv.getRange(2, 1, out.length, HP.length).setValues(out);

  // Sincronizar CANTIDAD en LOTES (match por LOTE+SKU+TALLA), conservando ESTADO/FECHA;
  // los combos nuevos (p. ej. tallas añadidas a mano) se añaden como PEDIDO.
  var shLot = ss.getSheetByName(SH.LOTES), HL = HEAD.LOTES, lastLo = shLot.getLastRow();
  var lotes = (lastLo >= 2) ? shLot.getRange(2, 1, lastLo - 1, HL.length).getValues() : [];
  var idxLote = {};
  lotes.forEach(function (r, i) { idxLote[String(r[HL.indexOf('LOTE')]) + '||' + r[HL.indexOf('SKU')] + '||' + r[HL.indexOf('TALLA')]] = i; });
  var nuevos = [];
  orden.forEach(function (k) {
    var a = agg[k], lk = a.lote + '||' + a.sku + '||' + a.talla;
    if (lk in idxLote) shLot.getRange(idxLote[lk] + 2, HL.indexOf('CANTIDAD') + 1).setValue(a.cantidad);
    else nuevos.push([a.lote, new Date(), a.producto, a.sku, a.talla, a.cantidad, 'PEDIDO', '']);
  });
  if (nuevos.length) shLot.getRange(shLot.getLastRow() + 1, 1, nuevos.length, HL.length).setValues(nuevos);

  regenerarResumenProveedor(ss);
  refrescarDashboard();
  var total = out.reduce(function (s, r) { return s + (Number(r[4]) || 0); }, 0);
  registrarLog(ss, 'PROVEEDOR_REBUILD', out.length + ' líneas · ' + total + ' uds desde LINEAS');
  ui().alert('PROVEEDOR reconstruida desde LINEAS_PEDIDO.\n\n' + out.length + ' líneas · ' + total +
    ' uds en total.\nRESUMEN_PROVEEDOR actualizado.' +
    (nuevos.length ? '\n\n(' + nuevos.length + ' combinación(es) de talla nuevas añadidas a LOTES.)' : ''));
}

function marcarLoteRecibidoSeleccion() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  if (sh.getName() !== SH.LOTES) { ui().alert('Ponte en la hoja LOTES y selecciona una fila del lote a recibir.'); return; }
  var fila = sh.getActiveCell().getRow();
  if (fila < 2) { ui().alert('Selecciona una fila de lote.'); return; }
  var H = HEAD.LOTES, lote = sh.getRange(fila, H.indexOf('LOTE') + 1).getValue();
  if (!lote) return;
  var last = sh.getLastRow(), vals = sh.getRange(2, 1, last - 1, H.length).getValues(), ahora = new Date();
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][H.indexOf('LOTE')]) === String(lote)) {
      sh.getRange(r + 2, H.indexOf('ESTADO') + 1).setValue('RECIBIDO');
      sh.getRange(r + 2, H.indexOf('FECHA_RECEPCION') + 1).setValue(ahora);
    }
  }
  var avisados = avisarPedidosListos(ss);
  registrarLog(ss, 'RECEPCION', lote + ' · avisados ' + avisados);
  refrescarDashboard();
  ui().alert('Lote ' + lote + ' marcado como RECIBIDO.\nPedidos avisados para recoger: ' + avisados);
}

function avisarPedidosListos(ss) {
  var lotesRec = {}, shLot = ss.getSheetByName(SH.LOTES), HL = HEAD.LOTES, lastLot = shLot.getLastRow();
  if (lastLot >= 2) shLot.getRange(2, 1, lastLot - 1, HL.length).getValues().forEach(function (r) {
    if (String(r[HL.indexOf('ESTADO')]) === 'RECIBIDO') lotesRec[String(r[HL.indexOf('LOTE')])] = true;
  });
  var shL = ss.getSheetByName(SH.LINEAS), H = HEAD.LINEAS, lastL = shL.getLastRow(), porPedido = {};
  if (lastL >= 2) shL.getRange(2, 1, lastL - 1, H.length).getValues().forEach(function (r) {
    var pid = String(r[H.indexOf('ID')]), lote = String(r[H.indexOf('LOTE')]);
    if (!porPedido[pid]) porPedido[pid] = { total: 0, listos: 0 };
    porPedido[pid].total++; if (lote && lotesRec[lote]) porPedido[pid].listos++;
  });
  var pedidos = indicePedidos(ss), cfg = leerConfig(), shP = ss.getSheetByName(SH.PEDIDOS), HP = HEAD.PEDIDOS, n = 0;
  for (var pid in porPedido) {
    var p = pedidos[pid]; if (!p || p.estado !== 'ENVIADO_PROVEEDOR') continue;
    var c = porPedido[pid];
    if (c.total > 0 && c.listos === c.total) {
      shP.getRange(p.fila, HP.indexOf('ESTADO') + 1).setValue('LISTO_RECOGIDA');
      shP.getRange(p.fila, HP.indexOf('FECHA_LISTO') + 1).setValue(new Date());
      try { emailListoRecoger(p.email, p.id, p.nombre, lineasDePedido(ss, p.id), p.productos, p.aportacion, p.total, cfg, p.site); n++; }
      catch (e) { registrarLog(ss, 'EMAIL_ERROR', 'listo ' + p.id + ': ' + e); }
    }
  }
  return n;
}

/* ===========================  ACCIONES MANUALES  ======================== */

function confirmarPagoSeleccion() {
  var r = pedidoSeleccionado(); if (!r) return;
  var ok = marcarPedidoPagado(SpreadsheetApp.getActive(), r, leerConfig());
  registrarLog(SpreadsheetApp.getActive(), 'PAGO_MANUAL', r.id);
  refrescarDashboard();
  ui().alert('Pedido ' + r.id + ' → PAGO_CONCILIADO.\n' +
    (ok ? 'Email de confirmación enviado a ' + r.email + '.' : '⚠️ El email NO se pudo enviar. Revisa la hoja LOG (fila EMAIL_ERROR).'));
}

/* Confirmación en BLOQUE por lista de IDs. Pensada para la conciliación por
   NOMBRE + IMPORTE (cuando el extracto NO trae el código AIR26 y casas fuera con
   scripts/casador.py → lista de IDs casados). Lee los IDs de la hoja CONFIRMAR_LOTE
   (columna A) y marca cada pedido PAGO_CONCILIADO + email, con la MISMA lógica que
   conciliarBanco: idempotente (los ya pagados se saltan), reanudable (si para por
   tiempo o por cuota de Gmail, re-ejecutar continúa donde lo dejó). */
var SH_CONFIRMAR = 'CONFIRMAR_LOTE';

function confirmarPagosPorLista() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SH_CONFIRMAR);
  if (!sh) {
    sh = ss.insertSheet(SH_CONFIRMAR);
    sh.getRange(1, 1, 1, 2).setValues([['ID', 'RESULTADO']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 180);
    ui().alert('He creado la hoja "' + SH_CONFIRMAR + '".\n\n' +
      'Pega en la columna A (debajo de "ID") los IDs de pedido a confirmar, uno por fila ' +
      '(p. ej. AIR26-00001). Puedes pegar el .txt de casados tal cual. Deja la columna B ' +
      'vacía: la rellena el script.\n\nLuego vuelve a pulsar "✅ Confirmar pagos por lista".');
    return;
  }
  var last = sh.getLastRow();
  if (last < 2) { ui().alert('La hoja "' + SH_CONFIRMAR + '" no tiene IDs. Pega los IDs en la columna A.'); return; }

  var vals = sh.getRange(2, 1, last - 1, 2).getValues();   // [ID, RESULTADO]
  var YA = { 'CONFIRMADO': 1, 'CONFIRMADO_SIN_EMAIL': 1, 'YA_PAGADO': 1 };
  var pendientes = 0;
  for (var i = 0; i < vals.length; i++) {
    var idp = String(vals[i][0]).trim();
    if (idp && !YA[String(vals[i][1]).trim().toUpperCase()]) pendientes++;
  }
  if (!pendientes) { ui().alert('No hay pedidos pendientes de confirmar en "' + SH_CONFIRMAR + '" (todos tienen ya un resultado).'); return; }

  var usaGmail = !PropertiesService.getScriptProperties().getProperty('EMAIL_API_KEY');
  var avisoQuota = usaGmail
    ? '\n\n⚠️ Sin proveedor (Brevo): los emails salen por Gmail, con límite de ~100/día. Si hay más, el script parará al agotar la cuota; continúa mañana re-ejecutando (es idempotente), o configura Brevo (📮).'
    : '';
  var resp = ui().alert('Confirmar pagos por lista',
    'Se van a CONFIRMAR ' + pendientes + ' pedido(s): pasan a PAGO_CONCILIADO y se envía el email ' +
    'de confirmación a cada uno.\n\nEsto NO se puede deshacer (los emails se envían).' + avisoQuota +
    '\n\n¿Continúas?', ui().ButtonSet.YES_NO);
  if (resp !== ui().Button.YES) return;

  var pedidos = indicePedidos(ss), cfg = leerConfig();
  var t0 = Date.now(), MAX_MS = 5 * 60 * 1000;   // margen bajo el límite de 6 min de Apps Script
  var conc = 0, ya = 0, noenc = 0, cad = 0, sinmail = 0, interrumpido = '';

  for (var r = 0; r < vals.length; r++) {
    var id = String(vals[r][0]).trim();
    if (!id || YA[String(vals[r][1]).trim().toUpperCase()]) continue;

    if (Date.now() - t0 > MAX_MS) { interrumpido = 'tiempo'; break; }
    if (usaGmail && MailApp.getRemainingDailyQuota() <= 0) { interrumpido = 'cuota'; break; }

    var p = pedidos[id];
    if (!p) { sh.getRange(r + 2, 2).setValue('NO_ENCONTRADO'); noenc++; continue; }
    if (ESTADOS_PAGADOS.indexOf(p.estado) >= 0) { sh.getRange(r + 2, 2).setValue('YA_PAGADO'); ya++; continue; }
    if (p.estado === 'CADUCADO') { sh.getRange(r + 2, 2).setValue('CADUCADO_OMITIDO'); cad++; continue; }

    var ok = marcarPedidoPagado(ss, p, cfg);
    pedidos[id].estado = 'PAGO_CONCILIADO';        // si el ID se repite en la lista, no re-dispara
    sh.getRange(r + 2, 2).setValue(ok ? 'CONFIRMADO' : 'CONFIRMADO_SIN_EMAIL');
    if (ok) conc++; else sinmail++;
  }

  registrarLog(ss, 'CONFIRMAR_LOTE', 'confirmados ' + conc + ' · sin email ' + sinmail +
    ' · ya ' + ya + ' · caducados ' + cad + ' · no encontrados ' + noenc +
    (interrumpido ? ' · PARADO(' + interrumpido + ')' : ''));
  refrescarDashboard();

  var msg = 'Confirmación por lista terminada.\n\n' +
    '✅ Confirmados (email enviado): ' + conc + '\n' +
    (sinmail ? '⚠️ Confirmados SIN email (revisa la hoja LOG): ' + sinmail + '\n' : '') +
    'Ya estaban pagados: ' + ya + '\n' +
    (cad ? 'Caducados (omitidos, revísalos a mano): ' + cad + '\n' : '') +
    (noenc ? 'IDs no encontrados: ' + noenc + '\n' : '');
  if (interrumpido === 'tiempo') msg += '\n⏱️ Parado por el límite de tiempo de Apps Script. Vuelve a pulsar para continuar donde lo dejó.';
  if (interrumpido === 'cuota') msg += '\n📭 Parado: cuota diaria de Gmail agotada. Continúa mañana re-ejecutando, o configura Brevo (📮).';
  ui().alert(msg);
}

/* Aviso "pedido revisado por duplicados / en proceso" en BLOQUE. Para los clientes
   cuya conciliación se resolvió A MANO (pedidos duplicados). Lee los IDs de la hoja
   REVISADOS_LOTE (columna A) y envía a cada uno su email con el ID como número de
   recogida, desde la cuenta del backend (enfadadosconairbus.contacto@gmail.com).
   Idempotente y reanudable como confirmarPagosPorLista(); NO cambia el estado del
   pedido, solo manda el aviso. */
var SH_REVISADOS = 'REVISADOS_LOTE';

function avisarPedidosRevisadosPorLista() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SH_REVISADOS);
  if (!sh) {
    sh = ss.insertSheet(SH_REVISADOS);
    sh.getRange(1, 1, 1, 2).setValues([['ID', 'RESULTADO']]).setFontWeight('bold');
    sh.setFrozenRows(1); sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 200);
    ui().alert('He creado la hoja "' + SH_REVISADOS + '".\n\n' +
      'Pega en la columna A los IDs de los pedidos revisados a mano (uno por fila). ' +
      'Deja la columna B vacía.\n\nLuego vuelve a pulsar "📨 Avisar pedidos revisados (lista)".');
    return;
  }
  var last = sh.getLastRow();
  if (last < 2) { ui().alert('La hoja "' + SH_REVISADOS + '" no tiene IDs. Pega los IDs en la columna A.'); return; }

  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  var YA = { 'AVISADO': 1 };
  var pend = 0;
  for (var i = 0; i < vals.length; i++) {
    var idp = String(vals[i][0]).trim();
    if (idp && !YA[String(vals[i][1]).trim().toUpperCase()]) pend++;
  }
  if (!pend) { ui().alert('No hay pedidos pendientes de avisar en "' + SH_REVISADOS + '" (todos tienen ya un resultado).'); return; }

  var usaGmail = !PropertiesService.getScriptProperties().getProperty('EMAIL_API_KEY');
  var avisoQuota = usaGmail
    ? '\n\n⚠️ Sin Brevo: salen por Gmail (límite ~100/día). Si hay más, para al agotar la cuota; re-ejecuta para continuar (es idempotente).'
    : '';
  var resp = ui().alert('Avisar pedidos revisados',
    'Se van a ENVIAR ' + pend + ' email(s) de "pedido revisado / en proceso", con el ID como número de recogida.\n\n' +
    'NO cambia el estado de los pedidos, solo envía el aviso. Los emails NO se pueden deshacer.' + avisoQuota +
    '\n\n¿Continúas?', ui().ButtonSet.YES_NO);
  if (resp !== ui().Button.YES) return;

  var pedidos = indicePedidos(ss), cfg = leerConfig();
  var t0 = Date.now(), MAX_MS = 5 * 60 * 1000;
  var env = 0, noenc = 0, sinmail = 0, err = 0, interrumpido = '';

  for (var r = 0; r < vals.length; r++) {
    var id = String(vals[r][0]).trim();
    if (!id || YA[String(vals[r][1]).trim().toUpperCase()]) continue;
    if (Date.now() - t0 > MAX_MS) { interrumpido = 'tiempo'; break; }
    if (usaGmail && MailApp.getRemainingDailyQuota() <= 0) { interrumpido = 'cuota'; break; }

    var p = pedidos[id];
    if (!p) { sh.getRange(r + 2, 2).setValue('NO_ENCONTRADO'); noenc++; continue; }
    if (!p.email) { sh.getRange(r + 2, 2).setValue('SIN_EMAIL'); sinmail++; continue; }
    try {
      emailPedidoRevisado(p.email, p.id, p.nombre, cfg);
      sh.getRange(r + 2, 2).setValue('AVISADO'); env++;
    } catch (e) {
      sh.getRange(r + 2, 2).setValue('ERROR');
      registrarLog(ss, 'EMAIL_ERROR', 'revisado ' + id + ': ' + e); err++;
    }
  }

  registrarLog(ss, 'AVISO_REVISADOS', 'avisados ' + env + ' · sin email ' + sinmail +
    ' · no encontrados ' + noenc + ' · error ' + err + (interrumpido ? ' · PARADO(' + interrumpido + ')' : ''));
  var msg = 'Avisos enviados: ' + env + '\n' +
    (sinmail ? 'Pedidos sin email: ' + sinmail + '\n' : '') +
    (noenc ? 'IDs no encontrados: ' + noenc + '\n' : '') +
    (err ? '⚠️ Errores de envío (revisa la hoja LOG): ' + err + '\n' : '');
  if (interrumpido === 'tiempo') msg += '\n⏱️ Parado por el límite de tiempo. Vuelve a pulsar para continuar.';
  if (interrumpido === 'cuota') msg += '\n📭 Cuota diaria de Gmail agotada. Continúa mañana, o configura Brevo (📮).';
  ui().alert(msg);
}

// Envía los 3 emails de ejemplo a EMAIL_CONTACTO para revisar el diseño sin recorrer el flujo.
function enviarEmailsPrueba() {
  var cfg = leerConfig(), to = cfg.EMAIL_CONTACTO;
  if (!to) { ui().alert('Pon EMAIL_CONTACTO en la hoja CONFIG.'); return; }
  var lineas = [{ producto: 'Camiseta', sku: 'CAMISETA-M', talla: 'M', cantidad: 2 },
                { producto: 'Camiseta', sku: 'CAMISETA-XL', talla: 'XL', cantidad: 1 }];
  try {
    emailPedidoRecibido(to, 'AIR26-PRUEBA', 'Prueba', lineas, 30, 10, 40, cfg);
    emailPagoConfirmado(to, 'AIR26-PRUEBA', 'Prueba', lineas, 30, 10, 40, cfg, 'San Pablo');
    emailListoRecoger(to, 'AIR26-PRUEBA', 'Prueba', lineas, 30, 10, 40, cfg, 'San Pablo');
    ui().alert('3 emails de prueba enviados a ' + to + '\n(recibido, confirmado y listo).');
  } catch (e) { ui().alert('No se pudieron enviar los emails de prueba:\n' + e); }
}

// Vacía los datos de prueba. Doble protección: interruptor MODO_PRUEBAS + confirmación escrita.
function resetearPruebas() {
  var ss = SpreadsheetApp.getActive(), cfgSh = ss.getSheetByName(SH.CONFIG), cfg = leerConfig();

  // Protección 1: interruptor MODO_PRUEBAS. Si no existe, lo crea en SI y pide volver a pulsar.
  if (cfg.MODO_PRUEBAS === undefined || cfg.MODO_PRUEBAS === '') {
    cfgSh.appendRow(['MODO_PRUEBAS', 'SI']);
    ui().alert('He añadido "MODO_PRUEBAS = SI" a CONFIG.\n\nMientras esté en SI podrás resetear.\nCuando salgas a producción, ponlo en NO para bloquear el borrado.\n\nVuelve a pulsar el botón para borrar ahora.');
    return;
  }
  if (String(cfg.MODO_PRUEBAS).trim().toUpperCase() !== 'SI') {
    ui().alert('🔒 RESETEO BLOQUEADO\n\nEn CONFIG, MODO_PRUEBAS = "' + cfg.MODO_PRUEBAS + '" (no es "SI").\nEsto protege los pedidos reales.\n\nSi de verdad quieres borrar, pon MODO_PRUEBAS = SI en CONFIG y vuelve a pulsar.');
    return;
  }

  // Protección 2: confirmación escrita.
  var resp = ui().prompt('⚠️ BORRAR DATOS DE PRUEBA',
    'Se vaciarán PEDIDOS, LINEAS_PEDIDO, MOVIMIENTOS_BANCO, LOTES, PROVEEDOR y LOG, y el contador AIR26 volverá a 0.\nNo se tocan CONFIG ni CATALOGO.\n\nEscribe BORRAR (mayúsculas) para confirmar:',
    ui().ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui().Button.OK || String(resp.getResponseText()).trim().toUpperCase() !== 'BORRAR') {
    ui().alert('Cancelado. No se ha borrado nada.'); return;
  }

  [SH.PEDIDOS, SH.LINEAS, SH.BANCO, SH.LOTES, SH.PROVEEDOR, SH.LOG].forEach(function (n) {
    var sh = ss.getSheetByName(n); if (!sh) return;
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, Math.max(1, sh.getLastColumn())).clearContent();
  });
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ULTIMO_NUM', '0');
  props.setProperty('ULTIMO_LOTE', '0');
  actualizarDatosPanel(ss); refrescarDashboard();
  ui().alert('✅ Datos de prueba borrados. Contador AIR26 reiniciado.\n\n⚠️ Antes de compartir la URL, pon MODO_PRUEBAS = NO en CONFIG para bloquear futuros borrados.');
}

function marcarEntregadoSeleccion() {
  var r = pedidoSeleccionado(); if (!r) return;
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SH.PEDIDOS), H = HEAD.PEDIDOS;
  sh.getRange(r.fila, H.indexOf('ESTADO') + 1).setValue('ENTREGADO');
  sh.getRange(r.fila, H.indexOf('FECHA_ENTREGADO') + 1).setValue(new Date());
  registrarLog(ss, 'ENTREGA', r.id); refrescarDashboard();
  ui().alert('Pedido ' + r.id + ' → ENTREGADO.');
}

// Libera pedidos TARDÍOS retenidos: RETENIDO → LIBERADO, para que entren al proveedor
// en la próxima "📦 Generar pedido a proveedor". Úsalo SOLO tras confirmar con el
// proveedor que llegan a tiempo. Trabaja sobre la SELECCIÓN (una o varias filas) de
// la hoja PEDIDOS, así que puedes marcar varios pedidos a la vez. No toca el ESTADO
// ni envía emails: solo levanta la retención.
function liberarPedidosTardiosSeleccion() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  if (sh.getName() !== SH.PEDIDOS) { ui().alert('Ponte en la hoja PEDIDOS y selecciona la(s) fila(s) de los pedidos tardíos a liberar.'); return; }
  var H = HEAD.PEDIDOS, colT = H.indexOf('TARDIO') + 1, colId = H.indexOf('ID') + 1;
  var rango = sh.getActiveRange();
  var fila0 = rango.getRow(), n = rango.getNumRows();
  if (fila0 < 2) { fila0 = 2; n = Math.max(0, sh.getLastRow() - 1); }   // si tocó la cabecera, toma todo
  if (n <= 0) { ui().alert('Selecciona al menos una fila de pedido.'); return; }

  var liberados = [], yaLib = 0, noTardios = 0;
  for (var i = 0; i < n; i++) {
    var f = fila0 + i;
    if (f > sh.getLastRow()) break;
    var val = String(sh.getRange(f, colT).getValue() || '').trim().toUpperCase();
    if (val === TARDIO_RETENIDO) {
      sh.getRange(f, colT).setValue(TARDIO_LIBERADO);
      liberados.push(String(sh.getRange(f, colId).getValue() || ''));
    } else if (val === TARDIO_LIBERADO) yaLib++;
    else noTardios++;
  }

  if (!liberados.length) {
    ui().alert('No he liberado nada.\n\n' +
      (yaLib ? yaLib + ' pedido(s) ya estaban LIBERADOS.\n' : '') +
      (noTardios ? noTardios + ' fila(s) no eran pedidos tardíos (RETENIDO).\n' : '') +
      '\nSelecciona en PEDIDOS las filas con TARDIO = RETENIDO.');
    return;
  }
  registrarLog(ss, 'TARDIO_LIBERADO', liberados.length + ' liberados: ' + liberados.join(', '));
  refrescarDashboard();
  ui().alert('✅ Liberados ' + liberados.length + ' pedido(s) tardío(s):\n\n' + liberados.join('\n') +
    '\n\nEntrarán a producción en la próxima "📦 Generar pedido a proveedor".' +
    (yaLib ? '\n\n(' + yaLib + ' ya estaban liberados.)' : '') +
    (noTardios ? '\n(' + noTardios + ' fila(s) no eran tardías.)' : ''));
}
function caducarPendientes() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SH.PEDIDOS), H = HEAD.PEDIDOS, last = sh.getLastRow();
  if (last < 2) return;
  var vals = sh.getRange(2, 1, last - 1, H.length).getValues(), ahora = new Date(), n = 0;
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][H.indexOf('ESTADO')]) === 'PENDIENTE_PAGO') {
      var cad = vals[r][H.indexOf('CADUCA')];
      if (cad instanceof Date && cad < ahora) { sh.getRange(r + 2, H.indexOf('ESTADO') + 1).setValue('CADUCADO'); n++; }
    }
  }
  registrarLog(ss, 'CADUCIDAD', n + ' caducados'); refrescarDashboard();
  ui().alert(n + ' pedido(s) marcados como CADUCADO.');
}
function pedidoSeleccionado() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  if (sh.getName() !== SH.PEDIDOS) { ui().alert('Ponte en la hoja PEDIDOS y selecciona la fila del pedido.'); return null; }
  var fila = sh.getActiveCell().getRow();
  if (fila < 2) { ui().alert('Selecciona una fila de pedido.'); return null; }
  return pedidoDeFila(ss, fila);
}

/* ===========================  EXPORTAR A EXCEL  ========================= */

function exportarExcel() {
  var ss = SpreadsheetApp.getActive();
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) { ui().alert('No se pudo exportar (código ' + resp.getResponseCode() + ').'); return; }
  var it = DriveApp.getFoldersByName('Tienda Airbus - Export');
  var carpeta = it.hasNext() ? it.next() : DriveApp.createFolder('Tienda Airbus - Export');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
  var file = carpeta.createFile(resp.getBlob().setName('Tienda-Airbus-' + stamp + '.xlsx'));
  registrarLog(ss, 'EXPORT', file.getName());
  ui().alert('Excel generado:\n\n' + file.getName() + '\n\nCarpeta "Tienda Airbus - Export" en tu Drive.\n' + file.getUrl());
}

/* ===========================  A2 · BACKUP DIARIO FUERA DE GOOGLE  ========= */
/* Envía el libro completo (.xlsx) por email a una dirección de OTRO proveedor
   (NO un @gmail de esta cuenta), para que un bloqueo de Google no borre los datos.
   El destino se lee de CONFIG (fila BACKUP_EMAIL); admite varios separados por comas.
   Un solo email al día: irrelevante para las cuotas de Gmail. */

function backupDiario() {
  var ss = SpreadsheetApp.getActive();
  var cfg = leerConfig();
  var dest = String(cfg.BACKUP_EMAIL || '').trim();
  if (!dest) { registrarLog(ss, 'BACKUP_OMITIDO', 'Falta la fila BACKUP_EMAIL en CONFIG'); return; }

  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) { registrarLog(ss, 'BACKUP_ERROR', 'export ' + resp.getResponseCode()); return; }

  var tz = Session.getScriptTimeZone();
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var stampFile = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmm');
  var blob = resp.getBlob().setName('Backup-Tienda-Airbus-' + stampFile + '.xlsx');

  var shP = ss.getSheetByName(SH.PEDIDOS);
  var nPedidos = shP ? Math.max(0, shP.getLastRow() - 1) : 0;

  MailApp.sendEmail({
    to: dest,
    subject: 'Backup tienda · ' + stamp + ' · ' + nPedidos + ' pedidos',
    body: 'Copia de seguridad automática del libro de la tienda (Caja de Resistencia).\n\n'
        + 'Fecha: ' + stamp + '\n'
        + 'Pedidos en el libro: ' + nPedidos + '\n\n'
        + 'Guarda este correo con su adjunto: es tu copia FUERA de la cuenta de Google.\n'
        + 'Si un día dejas de recibir este backup, revisa el disparador o el estado de la cuenta.',
    attachments: [blob],
    name: 'Backup Tienda Airbus'
  });
  registrarLog(ss, 'BACKUP_OK', dest + ' · ' + nPedidos + ' pedidos');
}

/* Crea (o recrea) el disparador diario de backupDiario, ~3:00. Un clic desde el menú. */
function programarBackupDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupDiario').timeBased().everyDays(1).atHour(3).create();
  var dest = String(leerConfig().BACKUP_EMAIL || '').trim();
  ui().alert('✅ Backup diario programado (~3:00).\n\n'
    + (dest ? ('Se enviará a: ' + dest) : '⚠️ Falta la fila BACKUP_EMAIL en CONFIG: añádela con un correo que NO sea de esta cuenta de Google.'));
}

/* ===========================  DASHBOARD (datos + gráficos)  ============= */

function refrescarDashboard() {
  var ss = SpreadsheetApp.getActive();
  var dash = ss.getSheetByName(SH.DASH); if (dash) dash.getRange('B2').setValue(new Date());
  actualizarDatosPanel(ss);   // rellena _PANEL_DATA (fuente de los gráficos)
}

// Cuenta pedidos por estado y camisetas conciliadas por talla → hoja oculta.
function actualizarDatosPanel(ss) {
  var data = ss.getSheetByName(SH.DATA); if (!data) return;

  // Distribución por estado.
  var pedidos = indicePedidos(ss), porEstado = {};
  ESTADOS.forEach(function (e) { porEstado[e] = 0; });
  for (var id in pedidos) { var e = pedidos[id].estado; if (porEstado[e] === undefined) porEstado[e] = 0; porEstado[e]++; }
  var filasEstado = ESTADOS.filter(function (e) { return porEstado[e] > 0; }).map(function (e) { return [e, porEstado[e]]; });
  if (!filasEstado.length) filasEstado = [['(sin pedidos)', 0]];

  // Camisetas conciliadas (pagadas) por talla.
  var pagados = {}; for (var id2 in pedidos) if (ESTADOS_PAGADOS.indexOf(pedidos[id2].estado) >= 0) pagados[id2] = true;
  var shL = ss.getSheetByName(SH.LINEAS), H = HEAD.LINEAS, lastL = shL.getLastRow(), porTalla = {};
  if (lastL >= 2) shL.getRange(2, 1, lastL - 1, H.length).getValues().forEach(function (r) {
    if (!pagados[String(r[H.indexOf('ID')])]) return;
    var t = String(r[H.indexOf('TALLA')]); porTalla[t] = (porTalla[t] || 0) + (Number(r[H.indexOf('CANTIDAD')]) || 0);
  });
  var ordenTallas = ['XS','S','M','L','XL','2XL','3XL','4XL','5XL'];
  var filasTalla = ordenTallas.filter(function (t) { return porTalla[t]; }).map(function (t) { return [t, porTalla[t]]; });
  if (!filasTalla.length) filasTalla = [['(sin datos)', 0]];

  data.clearContents();
  data.getRange(1, 1, 1, 2).setValues([['ESTADO', 'PEDIDOS']]);
  data.getRange(2, 1, filasEstado.length, 2).setValues(filasEstado);
  data.getRange(1, 4, 1, 2).setValues([['TALLA', 'UDS']]);
  data.getRange(2, 4, filasTalla.length, 2).setValues(filasTalla);
}

/* ===========================  TOKEN  =================================== */

function tokenValido(t) { var real = PropertiesService.getScriptProperties().getProperty('BACKEND_TOKEN'); return real && t && String(t) === String(real); }
function asegurarToken() {
  var props = PropertiesService.getScriptProperties(), t = props.getProperty('BACKEND_TOKEN');
  if (!t) { t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); props.setProperty('BACKEND_TOKEN', t); }
  return t;
}
function mostrarToken() { ui().alert('TOKEN backend (solo para el Cloudflare Worker):\n\n' + asegurarToken() + '\n\nNo lo pongas en GitHub ni en config.js.'); }

// Guarda la API key del proveedor de email (Brevo) en Script Properties.
function configurarEmailProveedor() {
  var u = ui();
  var resp = u.prompt('Proveedor de email (Brevo)',
    'Pega tu API key de Brevo (empieza por "xkeysib-").\n\nDéjalo vacío + Aceptar para BORRAR la key y volver a Gmail (100/día).',
    u.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== u.Button.OK) return;
  var key = String(resp.getResponseText()).trim();
  var props = PropertiesService.getScriptProperties();
  if (!key) { props.deleteProperty('EMAIL_API_KEY'); u.alert('Key borrada. Los emails saldrán por Gmail (límite 100/día).'); return; }
  props.setProperty('EMAIL_API_KEY', key);
  u.alert('API key guardada. Los emails saldrán por el proveedor.\n\nComprueba con "✉️ Enviar emails de prueba".\nOJO: el remitente (CONFIG → EMAIL_REMITENTE) debe estar VERIFICADO en Brevo.');
}

/* ===========================  EMAILS  =================================== */

// Envío de email a través de proveedor transaccional (Brevo). Si no hay API key
// configurada, cae a Gmail/MailApp (solo válido para bajo volumen / pruebas).
// El límite de volumen pasa a ser el del proveedor: se evita el bloqueo de Gmail.
function enviarEmail(to, subject, html) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('EMAIL_API_KEY');
  var cfg = leerConfig();
  var fromEmail = cfg.EMAIL_REMITENTE || cfg.EMAIL_CONTACTO || 'enfadadosconairbus.contacto@gmail.com';
  var fromName = cfg.EMAIL_REMITENTE_NOMBRE || 'Caja de Resistencia · Huelga Airbus';

  if (!key) {  // sin proveedor: Gmail (100/día, solo pruebas o bajo volumen)
    MailApp.sendEmail({ to: to, name: fromName, subject: subject, htmlBody: html });
    return;
  }
  // Brevo (https://developers.brevo.com). Para SES u otro, cambia SOLO esta llamada.
  var resp = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'api-key': key, 'accept': 'application/json' },
    payload: JSON.stringify({
      sender: { email: fromEmail, name: fromName },
      to: [{ email: to }],
      subject: subject,
      htmlContent: html
    })
  });
  var code = resp.getResponseCode();
  if (code >= 300) throw new Error('Proveedor email HTTP ' + code + ': ' + resp.getContentText().slice(0, 300));
}

// Nombre del site tal y como se enseña al colaborador ("San Pablo", "Getafe"...).
// El formulario ya lo guarda en Título, pero normalizamos por si llega en un solo
// caso (p. ej. "GETAFE" desde el sufijo de un lote, o "san pablo" de datos antiguos).
function siteNombre(site) {
  var s = String(site || '').trim();
  if (!s) return 'tu site de recogida';
  if (s === s.toUpperCase() || s === s.toLowerCase()) {
    s = s.toLowerCase().replace(/(^|\s)\S/g, function (c) { return c.toUpperCase(); });
  }
  return s;
}
// Clave de CONFIG para la dirección de envío de un site. Estable frente a
// mayúsculas/acentos/espacios: "San Pablo"→ENVIO_SAN_PABLO, "Cádiz"→ENVIO_CADIZ.
function envioKeySite(site) {
  var s = String(site || '').trim().toUpperCase();
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');   // fuera acentos
  s = s.replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');
  return 'ENVIO_' + s;
}
// Site DONDE SE RECOGE (regla de logística): Getafe e Illescas se recogen en
// Getafe; el resto en su propio site (envío por lotes al coordinador de logística).
// Vacío = pedidos antiguos sin SITE, que eran de Getafe.
function nombreRecogida(site) {
  var s = String(site || '').trim().toLowerCase();
  if (!s || s === 'getafe' || s === 'illescas') return 'Getafe';
  return siteNombre(site);
}
// Igual, pero para Getafe/Illescas devuelve la dirección completa de CONFIG (RECOGIDA).
function lugarRecogida(cfg, site) {
  var s = String(site || '').trim().toLowerCase();
  if (!s || s === 'getafe' || s === 'illescas') return cfg.RECOGIDA || 'Getafe';
  return siteNombre(site);
}

function emailPedidoRecibido(email, id, nombre, lineas, productos, aportacion, total, cfg) {
  enviarEmail(email, 'Aportación ' + id + ' recibida · Caja de Resistencia',
    plantillaEmail('Aportación recibida', 'Hola ' + escapar(nombre) + ', hemos registrado tu aportación <strong>' + id + '</strong>. Realiza una transferencia por el importe exacto usando <strong>' + id + '</strong> como concepto. No hace falta enviar justificante: confirmamos con los movimientos reales de la cuenta. ¡Gracias por colaborar con la caja de resistencia!', id, lineas, productos, aportacion, total, cfg, 'PENDIENTE DE TRANSFERENCIA'));
}
function emailPagoConfirmado(email, id, nombre, lineas, productos, aportacion, total, cfg, site) {
  enviarEmail(email, 'Aportación ' + id + ' confirmada · Caja de Resistencia',
    plantillaEmail('Aportación confirmada', 'Hola ' + escapar(nombre) + ', tu transferencia ha quedado <strong>confirmada</strong>. Te avisaremos por email cuando tu camiseta esté lista para recoger en <strong>' + escapar(nombreRecogida(site)) + '</strong>. ¡Gracias por tu apoyo!', id, lineas, productos, aportacion, total, cfg, 'CONFIRMADA'));
}
function emailListoRecoger(email, id, nombre, lineas, productos, aportacion, total, cfg, site) {
  enviarEmail(email, 'Tu camiseta ' + id + ' está lista para recoger',
    plantillaEmail('Lista para recoger', 'Hola ' + escapar(nombre) + ', tu camiseta de la aportación <strong>' + id + '</strong> ya está disponible. Recógela en: <strong>' + escapar(lugarRecogida(cfg, site)) + '</strong>.', id, lineas, productos, aportacion, total, cfg, 'LISTO PARA RECOGER'));
}
function emailPedidoRevisado(email, id, nombre, cfg) {
  enviarEmail(email, 'Tu pedido ' + id + ' está en proceso · Caja de Resistencia',
    plantillaPedidoRevisado(nombre, id, cfg));
}
// Aviso a clientes cuyo pedido se concilió A MANO por duplicados. Misma estética que
// plantillaEmail (barra roja, cabecera navy, pill, disclaimer) pero SIN líneas de
// producto: solo el mensaje y el ID destacado como número de recogida.
function plantillaPedidoRevisado(nombre, id, cfg) {
  return '' +
  '<div style="font-family:Arial,Helvetica,sans-serif;background:#f7f4ee;padding:24px 12px">' +
  '<div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #e4ddce;border-radius:16px;overflow:hidden">' +
  '<div style="height:6px;background:#c0392b;line-height:6px;font-size:6px">&nbsp;</div>' +
  '<div style="background:#16233b;color:#f3ede1;padding:18px 22px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;font-size:15px">Plataforma Solidaria · Caja de Resistencia' +
  '<span style="display:block;font-size:11px;font-weight:600;color:#b7ad9c;letter-spacing:.12em;margin-top:3px">Huelga Airbus 2026 · Albacete · Cádiz · Getafe · Illescas · San Pablo · Tablada</span></div>' +
  '<div style="padding:24px 22px">' +
  '<span style="display:inline-block;background:#e6ecf5;color:#16233b;font-weight:700;text-transform:uppercase;letter-spacing:.08em;font-size:11px;padding:5px 12px;border-radius:999px">Pedido en proceso</span>' +
  '<h1 style="margin:12px 0;font-size:24px;color:#1a1d21;text-transform:uppercase;letter-spacing:-.01em">Hemos revisado tu pedido</h1>' +
  '<p style="color:#4b4740;font-size:14px;line-height:1.6;margin:0 0 16px">Hola <strong>' + escapar(nombre) + '</strong>, gracias por tu aportación a la Caja de Resistencia. Hemos <strong>revisado en detalle tu pedido</strong> porque detectamos la existencia de <strong>pedidos duplicados</strong> a tu nombre. Ya está todo comprobado y tu pedido queda <strong>en proceso</strong>.</p>' +
  '<div style="margin-top:6px;padding:18px 20px;background:#faf6ee;border:1px solid #e4ddce;border-left:4px solid #16233b;border-radius:12px">' +
  '<div style="font-size:11px;color:#6f6a60;text-transform:uppercase;letter-spacing:.08em;font-weight:700">Tu número de pedido para la recogida</div>' +
  '<div style="margin-top:6px;font-family:Consolas,monospace;color:#c0392b;font-size:26px;font-weight:800;letter-spacing:.02em">' + escapar(id) + '</div>' +
  '<div style="margin-top:8px;font-size:13px;color:#4b4740;line-height:1.55">Guarda este número: es el que tendrás que dar al recoger tu camiseta. Aunque hubiera más de un pedido a tu nombre, <strong>este es el que vale</strong>.</div>' +
  '</div>' +
  '<p style="color:#4b4740;font-size:14px;line-height:1.6;margin:18px 0 0">Te avisaremos por email cuando tu camiseta esté <strong>lista para recoger</strong>. ¡Gracias por tu apoyo y por la paciencia!</p>' +
  '<p style="color:#9a9387;font-size:11px;margin-top:22px;line-height:1.5">Aportación solidaria a la Caja de Resistencia (Sindicato Útil); la camiseta es un agradecimiento por tu colaboración. Página no oficial de Airbus. No se procesan pagos: la aportación se realiza por transferencia bancaria.</p>' +
  '</div></div></div>';
}
function plantillaEmail(titulo, intro, id, lineas, productos, aportacion, total, cfg, estado) {
  var pill = ({
    'PENDIENTE DE TRANSFERENCIA': ['#fbe9e7', '#9c2c20'],
    'CONFIRMADA':                 ['#e4f1e8', '#2e7d52'],
    'LISTO PARA RECOGER':         ['#e6ecf5', '#16233b']
  })[estado] || ['#efeadf', '#6f6a60'];

  var filas = lineas.map(function (l, i) {
    var bg = (i % 2) ? '#faf7f0' : '#ffffff';
    return '<tr>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e4ddce;background:' + bg + ';color:#1a1d21;font-size:14px">' + escapar(l.producto) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e4ddce;background:' + bg + ';color:#1a1d21;font-size:14px">' + escapar(l.talla) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e4ddce;background:' + bg + ';color:#1a1d21;font-size:14px;text-align:center">' + l.cantidad + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #e4ddce;background:' + bg + ';color:#1a1d21;font-size:14px;text-align:right;font-weight:700">' + eur(precioSku(l.sku) * l.cantidad) + '</td></tr>';
  }).join('');

  var transferencia = (estado === 'PENDIENTE DE TRANSFERENCIA') ?
    '<div style="margin-top:18px;padding:16px 18px;background:#faf6ee;border:1px solid #e4ddce;border-left:4px solid #c0392b;border-radius:12px">' +
    '<div style="font-size:11px;color:#6f6a60;text-transform:uppercase;letter-spacing:.08em;font-weight:700">Datos de la transferencia</div>' +
    '<div style="margin-top:8px;font-size:14px;color:#1a1d21">Beneficiario: <strong>' + escapar(cfg.BENEFICIARIO || '') + '</strong></div>' +
    '<div style="font-size:14px;color:#1a1d21">IBAN: <strong style="font-family:Consolas,monospace">' + escapar(cfg.IBAN || '') + '</strong></div>' +
    '<div style="font-size:14px;color:#1a1d21;margin-top:4px">Concepto obligatorio: <strong style="font-family:Consolas,monospace;color:#c0392b;font-size:16px">' + id + '</strong></div>' +
    '<div style="margin-top:6px;font-size:12px;color:#6f6a60">Si tu banco no admite el guion, puedes ponerlo sin él (' + escapar(String(id).replace(/-/g, '')) + ') o con un espacio. Lo importante es que aparezca <strong>' + escapar(String(id).replace('-', ' ').split(' ')[0]) + '</strong> y el número.</div>' +
    '<div style="margin-top:8px;font-size:12px;color:#6f6a60">Revisamos las transferencias una vez al día: la confirmación puede tardar hasta 24 h. No hace falta enviar justificante.</div>' +
    '</div>' : '';

  return '' +
  '<div style="font-family:Arial,Helvetica,sans-serif;background:#f7f4ee;padding:24px 12px">' +
  '<div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #e4ddce;border-radius:16px;overflow:hidden">' +
  '<div style="height:6px;background:#c0392b;line-height:6px;font-size:6px">&nbsp;</div>' +
  '<div style="background:#16233b;color:#f3ede1;padding:18px 22px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;font-size:15px">Plataforma Solidaria · Caja de Resistencia' +
  '<span style="display:block;font-size:11px;font-weight:600;color:#b7ad9c;letter-spacing:.12em;margin-top:3px">Huelga Airbus 2026 · Albacete · Cádiz · Getafe · Illescas · San Pablo · Tablada</span></div>' +
  '<div style="padding:24px 22px">' +
  '<span style="display:inline-block;background:' + pill[0] + ';color:' + pill[1] + ';font-weight:700;text-transform:uppercase;letter-spacing:.08em;font-size:11px;padding:5px 12px;border-radius:999px">' + escapar(estado) + '</span>' +
  '<h1 style="margin:12px 0;font-size:24px;color:#1a1d21;text-transform:uppercase;letter-spacing:-.01em">' + escapar(titulo) + '</h1>' +
  '<p style="color:#4b4740;font-size:14px;line-height:1.6;margin:0 0 16px">' + intro + '</p>' +
  (lineas.length ?
  '<table style="width:100%;border-collapse:collapse;margin:6px 0 14px;border:1px solid #e4ddce;border-radius:10px;overflow:hidden">' +
  '<thead><tr>' +
  '<th style="text-align:left;padding:9px 12px;background:#16233b;color:#f3ede1;font-size:10px;text-transform:uppercase;letter-spacing:.06em">Producto</th>' +
  '<th style="text-align:left;padding:9px 12px;background:#16233b;color:#f3ede1;font-size:10px;text-transform:uppercase;letter-spacing:.06em">Talla</th>' +
  '<th style="text-align:center;padding:9px 12px;background:#16233b;color:#f3ede1;font-size:10px;text-transform:uppercase;letter-spacing:.06em">Uds.</th>' +
  '<th style="text-align:right;padding:9px 12px;background:#16233b;color:#f3ede1;font-size:10px;text-transform:uppercase;letter-spacing:.06em">Subtotal</th>' +
  '</tr></thead><tbody>' + filas + '</tbody></table>' +
  '<div style="text-align:right;color:#6f6a60;font-size:14px">Camisetas: <strong style="color:#1a1d21">' + eur(productos) + '</strong></div>' : '') +
  (aportacion > 0 ? '<div style="text-align:right;color:#6f6a60;font-size:14px">Aportación a la Caja: <strong style="color:#1a1d21">' + eur(aportacion) + '</strong></div>' : '') +
  '<div style="text-align:right;font-size:22px;color:#c0392b;font-weight:800;margin-top:6px">TOTAL: ' + eur(total) + '</div>' +
  transferencia +
  '<p style="color:#9a9387;font-size:11px;margin-top:22px;line-height:1.5">Aportación solidaria a la Caja de Resistencia (Sindicato Útil); la camiseta es un agradecimiento por tu colaboración. Página no oficial de Airbus. No se procesan pagos: la aportación se realiza por transferencia bancaria.</p>' +
  '</div></div></div>';
}

/* ===========================  LECTURAS / ÍNDICES  ====================== */

function indicePedidos(ss) {
  var sh = ss.getSheetByName(SH.PEDIDOS), H = HEAD.PEDIDOS, out = {}, last = sh.getLastRow();
  if (last < 2) return out;
  var vals = sh.getRange(2, 1, last - 1, H.length).getValues();
  for (var r = 0; r < vals.length; r++) { var id = String(vals[r][H.indexOf('ID')]); if (id) out[id] = filaAObjeto(vals[r], r + 2); }
  return out;
}
function pedidoDeFila(ss, fila) { var sh = ss.getSheetByName(SH.PEDIDOS), H = HEAD.PEDIDOS; return filaAObjeto(sh.getRange(fila, 1, 1, H.length).getValues()[0], fila); }
function filaAObjeto(row, fila) {
  var H = HEAD.PEDIDOS, g = function (k) { return row[H.indexOf(k)]; };
  return { fila: fila, id: String(g('ID')), nombre: g('NOMBRE'), email: g('EMAIL'), unidades: Number(g('UNIDADES')) || 0,
    productos: Number(g('PRODUCTOS_EUR')) || 0, aportacion: Number(g('APORTACION_EUR')) || 0, total: Number(g('TOTAL_EUR')) || 0,
    estado: String(g('ESTADO')), site: String(g('SITE') || ''), tardio: String(g('TARDIO') || '').trim().toUpperCase() };
}
function buscarPorCRID(ss, crid) {
  var sh = ss.getSheetByName(SH.PEDIDOS), H = HEAD.PEDIDOS, last = sh.getLastRow(); if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, H.length).getValues();
  for (var r = 0; r < vals.length; r++) if (String(vals[r][H.indexOf('CLIENT_REQUEST_ID')]) === crid) return { id: vals[r][H.indexOf('ID')], total: vals[r][H.indexOf('TOTAL_EUR')] };
  return null;
}
// Reenvía SOLO las confirmaciones que fallaron: las que quedaron como
// EMAIL_ERROR 'confirmado <id>' en el LOG. Idempotente — al reenviar con éxito
// deja 'EMAIL_REENVIADO <id>' en el LOG y no lo vuelve a mandar. Respeta la cuota
// de Gmail y el límite de tiempo de Apps Script (re-ejecuta para continuar).
// Úsalo DESPUÉS de arreglar el canal de email (Brevo verificado o key puesta).
function reenviarConfirmacionesFallidas() {
  var ss = SpreadsheetApp.getActive();
  var shLog = ss.getSheetByName(SH.LOG); if (!shLog) { ui().alert('No hay hoja LOG.'); return; }
  var last = shLog.getLastRow(); if (last < 2) { ui().alert('El LOG está vacío.'); return; }
  var HL = HEAD.LOG, iTipo = HL.indexOf('TIPO'), iDet = HL.indexOf('DETALLE');
  var log = shLog.getRange(2, 1, last - 1, HL.length).getValues();

  var fallidos = {}, reenviados = {};
  log.forEach(function (r) {
    var tipo = String(r[iTipo] || ''), det = String(r[iDet] || '');
    if (tipo === 'EMAIL_ERROR') { var m = det.match(/^confirmado\s+(\S+?):/); if (m) fallidos[m[1]] = true; }
    else if (tipo === 'EMAIL_REENVIADO') { var id0 = det.trim().split(/\s+/)[0]; if (id0) reenviados[id0] = true; }
  });
  var ids = Object.keys(fallidos).filter(function (id) { return !reenviados[id]; });
  if (!ids.length) { ui().alert('No hay confirmaciones fallidas pendientes de reenviar.'); return; }

  var usaGmail = !PropertiesService.getScriptProperties().getProperty('EMAIL_API_KEY');
  var aviso = usaGmail ? '\n\n⚠️ Sin Brevo: salen por Gmail (~100/día). Si hay más, parará al agotar la cuota; re-ejecuta mañana (es idempotente).' : '';
  var resp = ui().alert('Reenviar confirmaciones',
    'Se reenviará el email de confirmación a ' + ids.length + ' pedido(s) cuyo correo había fallado.' + aviso +
    '\n\nHazlo solo cuando el canal de email esté arreglado. ¿Continúas?', ui().ButtonSet.YES_NO);
  if (resp !== ui().Button.YES) return;

  var pedidos = indicePedidos(ss), cfg = leerConfig();
  var t0 = Date.now(), MAX_MS = 5 * 60 * 1000;
  var ok = 0, fail = 0, omit = 0, parado = '';
  for (var i = 0; i < ids.length; i++) {
    if (Date.now() - t0 > MAX_MS) { parado = 'tiempo'; break; }
    if (usaGmail && MailApp.getRemainingDailyQuota() <= 0) { parado = 'cuota'; break; }
    var id = ids[i], p = pedidos[id];
    if (!p || ESTADOS_PAGADOS.indexOf(p.estado) < 0) { omit++; continue; }   // no existe / no pagado
    try {
      emailPagoConfirmado(p.email, p.id, p.nombre, lineasDePedido(ss, p.id), p.productos, p.aportacion, p.total, cfg, p.site);
      registrarLog(ss, 'EMAIL_REENVIADO', p.id);
      ok++;
    } catch (e) { registrarLog(ss, 'EMAIL_ERROR', 'reenvio ' + p.id + ': ' + e); fail++; }
  }
  refrescarDashboard();
  var msg = 'Reenvío terminado.\n\n✅ Enviados: ' + ok +
    (fail ? '\n❌ Fallaron otra vez: ' + fail + ' (mira el LOG)' : '') +
    (omit ? '\n• Omitidos (no pagados / no encontrados): ' + omit : '');
  if (parado === 'tiempo') msg += '\n\n⏱️ Parado por tiempo; vuelve a pulsar para continuar donde lo dejó.';
  if (parado === 'cuota') msg += '\n\n📭 Cuota de Gmail agotada; continúa mañana (o configura Brevo con 📮).';
  ui().alert(msg);
}

// Envía el email de "aportación confirmada" a los pedidos pagados que aún no están
// listos para recoger (PAGO_CONCILIADO y ENVIADO_PROVEEDOR — este último ya loteado).
// Para cuando las confirmaciones no llegaron en su momento. Idempotente: deja
// 'EMAIL_CONF_MANUAL <id>' en el LOG y no reenvía a ese pedido en futuras pasadas.
// Respeta cuota de Gmail y límite de tiempo (re-ejecuta para continuar).
// ANTES de usarlo: confirma con "✉️ Enviar emails de prueba" que el correo LLEGA (bandeja, no spam).
function enviarConfirmacionesPagoConciliado() {
  var ss = SpreadsheetApp.getActive();
  var pedidos = indicePedidos(ss), cfg = leerConfig();

  // Ya enviados por esta vía, para no duplicar en re-ejecuciones.
  var enviados = {};
  var shLog = ss.getSheetByName(SH.LOG);
  if (shLog && shLog.getLastRow() >= 2) {
    var HL = HEAD.LOG, il = shLog.getRange(2, 1, shLog.getLastRow() - 1, HL.length).getValues();
    var iT = HL.indexOf('TIPO'), iD = HL.indexOf('DETALLE');
    il.forEach(function (r) { if (String(r[iT]) === 'EMAIL_CONF_MANUAL') { var id0 = String(r[iD]).trim().split(/\s+/)[0]; if (id0) enviados[id0] = true; } });
  }

  // Pagados a los que aplica la confirmación de transferencia: los que aún no han
  // llegado a "listo para recoger". Incluye ENVIADO_PROVEEDOR (ya loteados).
  var ESTADOS_CONF = ['PAGO_CONCILIADO', 'ENVIADO_PROVEEDOR'];
  var ids = [];
  for (var id in pedidos) { if (ESTADOS_CONF.indexOf(pedidos[id].estado) >= 0 && !enviados[id]) ids.push(id); }
  if (!ids.length) { ui().alert('No hay pedidos pagados (PAGO_CONCILIADO / ENVIADO_PROVEEDOR) pendientes de confirmar por email.'); return; }

  var usaGmail = !PropertiesService.getScriptProperties().getProperty('EMAIL_API_KEY');
  var aviso = usaGmail ? '\n\n⚠️ Sin Brevo: por Gmail (~100/día). Re-ejecuta mañana para continuar (es idempotente).' : '';
  var resp = ui().alert('Enviar confirmación de pago',
    'Se enviará el email de "aportación confirmada" a ' + ids.length + ' pedido(s) pagados (PAGO_CONCILIADO y ENVIADO_PROVEEDOR).\n\n' +
    '⚠️ Comprueba ANTES que un email de prueba te llega a la BANDEJA (no spam).' + aviso + '\n\n¿Continúas?', ui().ButtonSet.YES_NO);
  if (resp !== ui().Button.YES) return;

  var t0 = Date.now(), MAX_MS = 5 * 60 * 1000, ok = 0, fail = 0, parado = '';
  for (var i = 0; i < ids.length; i++) {
    if (Date.now() - t0 > MAX_MS) { parado = 'tiempo'; break; }
    if (usaGmail && MailApp.getRemainingDailyQuota() <= 0) { parado = 'cuota'; break; }
    var p = pedidos[ids[i]];
    try {
      emailPagoConfirmado(p.email, p.id, p.nombre, lineasDePedido(ss, p.id), p.productos, p.aportacion, p.total, cfg, p.site);
      registrarLog(ss, 'EMAIL_CONF_MANUAL', p.id); ok++;
    } catch (e) { registrarLog(ss, 'EMAIL_ERROR', 'conf-manual ' + p.id + ': ' + e); fail++; }
  }
  refrescarDashboard();
  var msg = 'Envío terminado.\n\n✅ Enviados: ' + ok + (fail ? '\n❌ Fallaron: ' + fail + ' (mira el LOG)' : '');
  if (parado === 'tiempo') msg += '\n\n⏱️ Parado por tiempo; vuelve a pulsar para continuar.';
  if (parado === 'cuota') msg += '\n\n📭 Cuota de Gmail agotada; continúa mañana (o configura Brevo).';
  ui().alert(msg);
}

function lineasDePedido(ss, id) {
  var sh = ss.getSheetByName(SH.LINEAS), H = HEAD.LINEAS, last = sh.getLastRow(), out = []; if (last < 2) return out;
  sh.getRange(2, 1, last - 1, H.length).getValues().forEach(function (r) {
    if (String(r[H.indexOf('ID')]) === String(id)) out.push({ producto: r[H.indexOf('PRODUCTO')], sku: r[H.indexOf('SKU')], talla: r[H.indexOf('TALLA')], cantidad: Number(r[H.indexOf('CANTIDAD')]) || 0 });
  });
  return out;
}
var _catCache = null;
function leerCatalogo(ss) {
  var sh = ss.getSheetByName(SH.CATALOGO), out = {}, last = sh.getLastRow(); if (last < 2) return out;
  sh.getRange(2, 1, last - 1, HEAD.CATALOGO.length).getValues().forEach(function (row) {
    var sku = String(row[2]).trim(); if (!sku) return;
    out[sku] = { activo: row[0] === true || String(row[0]).toUpperCase() === 'TRUE', producto: row[1], sku: sku, talla: row[3], medidas: row[4], precio: Number(row[5]) || 0, coste: Number(row[6]) || 0, aporte: Number(row[7]) || 0 };
  });
  _catCache = out; return out;
}
function precioSku(sku) { if (!_catCache) leerCatalogo(SpreadsheetApp.getActive()); var it = _catCache[sku]; return it ? it.precio : 0; }
function leerConfig() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.CONFIG), out = {}; if (!sh) return out;
  var last = sh.getLastRow(); if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}
function registrarLog(ss, tipo, detalle) {
  try { var sh = ss.getSheetByName(SH.LOG); if (sh) sh.appendRow([new Date(), tipo, detalle]); } catch (e) {}
}

/* ===========================  IDs  ==================================== */

function siguienteId(cfg) {
  var props = PropertiesService.getScriptProperties(), n = Number(props.getProperty('ULTIMO_NUM') || '0') + 1;
  props.setProperty('ULTIMO_NUM', String(n));
  return (cfg.PREFIJO || 'AIR26') + '-' + pad(n, 5);
}
function nuevoLoteId(ss) {
  var props = PropertiesService.getScriptProperties(), n = Number(props.getProperty('ULTIMO_LOTE') || '0') + 1;
  props.setProperty('ULTIMO_LOTE', String(n));
  return 'LOTE-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + pad(n, 3);
}
// Detecta el código de pedido en el concepto de un movimiento bancario, tolerando
// las variantes que pone la gente cuando el banco no admite el guion: "AIR26-00123",
// "AIR26 00123", "AIR26_00123", "AIR2600123", "AIR26.00123", varios espacios, etc.
// Aplana el texto (quita todo lo no alfanumérico) y localiza el prefijo + el número real.
// OJO: hay que saltar los ceros de relleno ANTES de capturar el número, y NO confundir
// el "26" del prefijo con el número del pedido (bug de la versión anterior).
function detectarId(concepto, cfg) {
  var pref = String((cfg && cfg.PREFIJO) || 'AIR26').toUpperCase().replace(/[^A-Z0-9]/g, '');
  var flat = String(concepto).toUpperCase().replace(/[^A-Z0-9]/g, '');
  var m = flat.match(new RegExp(pref + '0*(\\d{1,6})'));
  return m ? pref + '-' + pad(Number(m[1]), 5) : '';
}
function pad(n, len) { var s = String(n); while (s.length < len) s = '0' + s; return s; }
// Site → sufijo legible para el ID del lote: mayúsculas, sin acentos ni espacios.
// "Cádiz" → "CADIZ", "San Pablo" → "SANPABLO".
function slugSite(s) {
  return String(s || 'SINSITE').toUpperCase()
    .replace(/[ÁÀÄÂ]/g,'A').replace(/[ÉÈËÊ]/g,'E').replace(/[ÍÌÏÎ]/g,'I').replace(/[ÓÒÖÔ]/g,'O').replace(/[ÚÙÜÛ]/g,'U').replace(/Ñ/g,'N')
    .replace(/[^A-Z0-9]+/g, '') || 'SINSITE';
}

/* ===========================  SETUP + FORMATO  ======================== */

function setupTiendaV4() {
  var ss = SpreadsheetApp.getActive();

  // Limpia el nombre antiguo 'BANCO' (versión previa) si está vacío.
  var viejo = ss.getSheetByName('BANCO');
  if (viejo && viejo.getLastRow() <= 1) ss.deleteSheet(viejo);

  hoja(ss, SH.CONFIG, ['CLAVE', 'VALOR']);
  hoja(ss, SH.CATALOGO, HEAD.CATALOGO);
  hoja(ss, SH.PEDIDOS, HEAD.PEDIDOS);
  hoja(ss, SH.LINEAS, HEAD.LINEAS);
  hoja(ss, SH.BANCO, HEAD.BANCO);
  hoja(ss, SH.LOTES, HEAD.LOTES);
  hoja(ss, SH.PROVEEDOR, HEAD.PROVEEDOR);
  hoja(ss, SH.LOG, HEAD.LOG);

  var cfg = ss.getSheetByName(SH.CONFIG);
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, 19, 2).setValues([
    ['BENEFICIARIO', 'Caja de Resistencia Huelga Airbus 2026 - Sindicato Útil'],
    ['IBAN', 'ESXX XXXX XXXX XXXX XXXX XXXX  [COMPLETAR ANTES DE PUBLICAR]'],
    ['EMAIL_CONTACTO', 'enfadadosconairbus.contacto@gmail.com'],
    ['RECOGIDA', 'Getafe - Factoría Airbus - Puerta Sur / Puerta Norte (Asamblea de trabajadores en Huelga)'],
    ['CADUCIDAD_HORAS', 12], ['MAX_UNIDADES', 20], ['PREFIJO', 'AIR26'],
    ['MODO_PRUEBAS', 'SI'],
    ['EMAIL_REMITENTE', 'enfadadosconairbus.contacto@gmail.com'],
    ['EMAIL_REMITENTE_NOMBRE', 'Caja de Resistencia · Huelga Airbus'],
    ['PROD_TODOS_SITES_DESDE', '2026-09-10'],
    // Fecha límite para producir a tiempo. Los pedidos creados a partir de este
    // instante se marcan TARDIO=RETENIDO y NO entran solos al proveedor (hay que
    // liberarlos a mano tras confirmar plazos). Debe ir en ISO con zona horaria y
    // coincidir con AVISO_FECHA_LIMITE de config.js (web).
    ['AVISO_FECHA_LIMITE', '2026-09-06T21:00:00+02:00'],
    // Direcciones de envío por site (el proveedor envía). RESUMEN_PROVEEDOR las replica.
    ['ENVIO_GETAFE', '[COMPLETAR dirección de envío · Getafe]'],
    ['ENVIO_ILLESCAS', '[COMPLETAR dirección de envío · Illescas]'],
    ['ENVIO_SAN_PABLO', '[COMPLETAR dirección de envío · San Pablo]'],
    ['ENVIO_TABLADA', '[COMPLETAR dirección de envío · Tablada]'],
    ['ENVIO_ALBACETE', '[COMPLETAR dirección de envío · Albacete]'],
    ['ENVIO_CADIZ', '[COMPLETAR dirección de envío · Cádiz]']
  ]);

  var cat = ss.getSheetByName(SH.CATALOGO);
  if (cat.getLastRow() < 2) {
    var tallas = [['XS','46×66'],['S','49×69'],['M','52×71'],['L','55×73'],['XL','58×75'],['2XL','62×77'],['3XL','66×79'],['4XL','70×81'],['5XL','74×83']];
    cat.getRange(2, 1, 9, HEAD.CATALOGO.length).setValues(tallas.map(function (t) { return [true, 'Camiseta', 'CAMISETA-' + t[0], t[0], t[1], 10, 5, 5]; }));
  }

  asegurarToken();
  reconstruirEstetica();  // formato + panel + gráficos + how_to
  ui().alert('Backend V4.2 preparado (con panel y gráficos).\n\n1) Revisa CONFIG (IBAN y beneficiario reales).\n2) Implementa como Aplicación web.\n3) Menú → 🔑 Mostrar TOKEN backend.');
}

// Reaplica todo el formato/estética sin tocar los datos (se puede correr cuando quieras).
function reconstruirEstetica() {
  var ss = SpreadsheetApp.getActive();
  hoja(ss, SH.DATA, ['ESTADO', 'PEDIDOS']);
  formatearHojasDatos(ss);
  formatoCondicionalEstados(ss);
  validaciones(ss);
  construirHowTo(ss);
  construirDashboard(ss);
  actualizarDatosPanel(ss);
  insertarGraficos(ss);
  ordenarPestanas(ss);
  ss.getSheetByName(SH.DATA).hideSheet();
}

function formatearHojasDatos(ss) {
  var mapa = [
    [SH.PEDIDOS, HEAD.PEDIDOS, { moneda: ['PRODUCTOS_EUR','APORTACION_EUR','TOTAL_EUR'], fecha: ['FECHA_PEDIDO','CADUCA','FECHA_CONFIRMADO','FECHA_LISTO','FECHA_ENTREGADO'] }],
    [SH.LINEAS, HEAD.LINEAS, { fecha: ['FECHA_PEDIDO'] }],
    [SH.CATALOGO, HEAD.CATALOGO, { moneda: ['PRECIO','COSTE','APORTE_CAJA'] }],
    [SH.BANCO, HEAD.BANCO, { moneda: ['IMPORTE'], fecha: ['FECHA','FECHA_CONCILIACION'] }],
    [SH.LOTES, HEAD.LOTES, { fecha: ['FECHA_GENERACION','FECHA_RECEPCION'] }],
    [SH.PROVEEDOR, HEAD.PROVEEDOR, {}],
    [SH.LOG, HEAD.LOG, { fecha: ['TIMESTAMP'] }]
  ];
  mapa.forEach(function (m) {
    var sh = ss.getSheetByName(m[0]); if (!sh) return;
    var H = m[1], fmt = m[2];
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setFontColor(COL.paper).setBackground(COL.ink).setVerticalAlignment('middle');
    sh.setRowHeight(1, 28);
    (fmt.moneda || []).forEach(function (c) { sh.getRange(2, H.indexOf(c) + 1, 5000, 1).setNumberFormat(FMT_EUR); });
    (fmt.fecha || []).forEach(function (c) { sh.getRange(2, H.indexOf(c) + 1, 5000, 1).setNumberFormat(FMT_FECHA); });
    sh.autoResizeColumns(1, H.length);
    for (var c = 1; c <= H.length; c++) { var w = sh.getColumnWidth(c); if (w < 90) sh.setColumnWidth(c, 90); if (w > 260) sh.setColumnWidth(c, 260); }
  });
}

function formatoCondicionalEstados(ss) {
  var pares = [
    ['PENDIENTE_PAGO', COL.tPeach], ['PAGO_CONCILIADO', COL.tGreen], ['ENVIADO_PROVEEDOR', COL.tBlue],
    ['RECIBIDO', COL.tBlue2], ['LISTO_RECOGIDA', COL.tYellow], ['ENTREGADO', COL.tGray], ['CADUCADO', COL.tRed]
  ];
  var sh = ss.getSheetByName(SH.PEDIDOS), col = HEAD.PEDIDOS.indexOf('ESTADO') + 1;
  var rango = sh.getRange(2, col, 5000, 1), reglas = [];
  pares.forEach(function (p) {
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(p[0]).setBackground(p[1]).setRanges([rango]).build());
  });
  // Columna TARDIO: RETENIDO en rojo (retenido de producción), LIBERADO en verde.
  var colT = HEAD.PEDIDOS.indexOf('TARDIO') + 1, rangoT = sh.getRange(2, colT, 5000, 1);
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('RETENIDO').setBackground(COL.tRed).setRanges([rangoT]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('LIBERADO').setBackground(COL.tGreen).setRanges([rangoT]).build());
  sh.setConditionalFormatRules(reglas);

  // Banco: REVISAR en rojo, conciliado en verde.
  var b = ss.getSheetByName(SH.BANCO), cb = HEAD.BANCO.indexOf('RESULTADO') + 1, rb = b.getRange(2, cb, 5000, 1);
  b.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('REVISAR').setBackground(COL.tRed).setRanges([rb]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PAGO_CONCILIADO').setBackground(COL.tGreen).setRanges([rb]).build()
  ]);
}

function validaciones(ss) {
  var pe = ss.getSheetByName(SH.PEDIDOS);
  pe.getRange(2, HEAD.PEDIDOS.indexOf('ESTADO') + 1, 5000, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(true).build());
  var cat = ss.getSheetByName(SH.CATALOGO);
  cat.getRange(2, 1, 100, 1).insertCheckboxes();  // ACTIVO
}

function construirHowTo(ss) {
  var sh = ss.getSheetByName(SH.HOWTO) || ss.insertSheet(SH.HOWTO, 0);
  sh.clear();
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), 20), Math.max(sh.getMaxColumns(), 8)).breakApart();
  sh.setHiddenGridlines(true);
  banner(sh, 'A1:H1', 'CAMISETAS SOLIDARIAS · CENTRO DE OPERACIONES', 16);
  sub(sh, 'A2:H2', 'Cómo operar, en qué orden y qué no tocar');
  var secciones = [
    ['🏦 IMPORTAR BANCO', 'Pega el extracto (FECHA · CONCEPTO · IMPORTE · REFERENCIA) en la hoja MOVIMIENTOS_BANCO.'],
    ['🔄 CONCILIAR', 'Menú 👕 → 🏦 Conciliar banco. Los matches exactos (código + importe) pasan a PAGO_CONCILIADO y reciben email.'],
    ['⚠️ REVISAR', 'Filtra RESULTADO = REVISAR. No fuerces matches dudosos por nombre; usa "Confirmar PAGO (manual)" solo si lo verificas.'],
    ['📦 CERRAR LOTE', 'Menú 👕 → 📦 Generar pedido a proveedor. Solo entran PAGO_CONCILIADO sin lote. Agrega por PRODUCTO+SKU+TALLA. Los pedidos TARDÍOS (columna TARDIO = RETENIDO) NO entran solos: quedan fuera hasta liberarlos.'],
    ['🕒 TARDÍOS', 'Pedidos hechos tras la fecha límite (CONFIG → AVISO_FECHA_LIMITE). Se cobran y concilian igual, pero se RETIENEN de producción. Confirma plazos con el proveedor → selecciona sus filas en PEDIDOS → 🕒 Liberar pedidos tardíos. Pasan a LIBERADO y ya entran al siguiente lote.'],
    ['📤 ENVIAR PROVEEDOR', 'Usa la hoja PROVEEDOR (resumen limpio por talla del lote). Exporta a Excel si lo necesitas.'],
    ['📥 RECIBIR', 'Cuando llegue la mercancía: hoja LOTES → selecciona el lote → 📥 Marcar lote recibido. Los pedidos completos pasan a LISTO_RECOGIDA y avisan por email.'],
    ['🤝 ENTREGAR', 'Al entregar en mano: hoja PEDIDOS → selecciona la fila → 🤝 Marcar ENTREGADO.']
  ];
  var r = 4;
  secciones.forEach(function (s) {
    sh.getRange(r, 1).setValue(s[0]).setFontWeight('bold').setFontColor(COL.orange).setFontSize(12);
    sh.getRange(r, 2, 1, 7).merge().setValue(s[1]).setWrap(true).setVerticalAlignment('middle').setFontColor(COL.ink);
    sh.setRowHeight(r, 34); r++;
  });
  r++;
  sh.getRange(r, 1, 1, 8).merge()
    .setValue('REGLA DE ORO · El Google Sheet es la fuente de verdad. Un pedido solo entra al proveedor cuando está en PAGO_CONCILIADO y sin lote. No conciliamos por nombre: código AIR26-XXXXX + importe exacto.')
    .setWrap(true).setBackground(COL.tYellow).setFontColor(COL.ink).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(r, 54);
  sh.setColumnWidth(1, 190); for (var c = 2; c <= 8; c++) sh.setColumnWidth(c, 150);
}

function construirDashboard(ss) {
  var sh = ss.getSheetByName(SH.DASH) || ss.insertSheet(SH.DASH, 1);
  sh.clear(); sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), 40), Math.max(sh.getMaxColumns(), 9)).breakApart();
  sh.setHiddenGridlines(true);
  for (var c = 1; c <= 9; c++) sh.setColumnWidth(c, c === 1 ? 24 : 150);

  banner(sh, 'A1:I1', 'DASHBOARD · CAMISETAS SOLIDARIAS AIRBUS 2026', 16);
  sh.getRange('A2').setValue('Actualizado:').setFontColor(COL.slate).setFontStyle('italic');
  sh.getRange('B2').setNumberFormat(FMT_FECHA);

  var P = SH.PEDIDOS, K = P + '!K2:K5000';
  var mask = '(' + K + '<>"PENDIENTE_PAGO")*(' + K + '<>"CADUCADO")*(' + K + '<>"")';
  var tiles = [
    ['PEDIDOS TOTALES', '=COUNTA(' + P + '!A2:A5000)', COL.tGray, false],
    ['PAGOS CONCILIADOS', '=SUMPRODUCT(' + mask + ')', COL.tGreen, false],
    ['CAMISETAS CONCILIADAS', '=SUMPRODUCT(' + mask + '*' + P + '!G2:G5000)', COL.tGreen2, false],
    ['CAJA GENERADA', '=SUMPRODUCT(' + mask + '*' + P + '!I2:I5000)+5*SUMPRODUCT(' + mask + '*' + P + '!G2:G5000)', COL.tYellow, true],
    ['PENDIENTES', '=COUNTIF(' + P + '!K2:K5000,"PENDIENTE_PAGO")', COL.tPeach, false],
    ['CADUCADOS', '=COUNTIF(' + P + '!K2:K5000,"CADUCADO")', COL.tRed, false],
    ['A REVISAR BANCO', '=COUNTIF(' + SH.BANCO + '!F2:F5000,"*REVISAR*")', COL.tRed, false],
    ['INGRESO CONCILIADO', '=SUMPRODUCT(' + mask + '*' + P + '!J2:J5000)', COL.tBlue, true]
  ];
  // 4 tarjetas por fila, cada una 2 columnas (B:C, D:E, F:G, H:I), 2 filas (label/valor).
  var startCols = [2, 4, 6, 8];
  for (var t = 0; t < tiles.length; t++) {
    var fila = t < 4 ? 4 : 7, ci = startCols[t % 4];
    tarjeta(sh, fila, ci, tiles[t][0], tiles[t][1], tiles[t][2], tiles[t][3]);
  }

  sh.getRange('B10:D10').merge().setValue('DISTRIBUCIÓN POR ESTADO').setFontWeight('bold').setFontColor(COL.ink);
  sh.getRange('F10:I10').merge().setValue('CAMISETAS CONCILIADAS POR TALLA').setFontWeight('bold').setFontColor(COL.ink);

  // Regla de oro
  sh.getRange('A26:I26').merge().setValue('REGLA DE ORO · Un pedido solo entra al proveedor cuando está en PAGO_CONCILIADO y sin lote. Conciliamos por código AIR26-XXXXX + importe exacto, nunca por nombre.')
    .setWrap(true).setBackground(COL.tYellow).setFontColor(COL.ink).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(26, 46);

  // Accesos rápidos
  sh.getRange('A28').setValue('ACCESOS RÁPIDOS').setFontWeight('bold').setFontColor(COL.orange);
  var links = [['⚙️ CONFIG', SH.CONFIG], ['🧾 PEDIDOS', SH.PEDIDOS], ['🏦 BANCO', SH.BANCO], ['📦 LOTES', SH.LOTES], ['🏭 PROVEEDOR', SH.PROVEEDOR], ['🧬 LÍNEAS', SH.LINEAS], ['🗒️ LOG', SH.LOG]];
  var cc = 2;
  links.forEach(function (lk) {
    var target = ss.getSheetByName(lk[1]); if (!target) return;
    sh.getRange(29, cc).setFormula('=HYPERLINK("#gid=' + target.getSheetId() + '","' + lk[0] + '")').setFontColor(COL.orangeDark);
    cc++;
  });

  // Leyenda
  sh.getRange('A31').setValue('LECTURA RÁPIDA').setFontWeight('bold').setFontColor(COL.orange);
  sh.getRange('A32:I32').merge().setValue('🟠 PENDIENTE_PAGO: creado, sin pago.   🟢 PAGO_CONCILIADO: puede entrar al lote.   🔵 ENVIADO_PROVEEDOR: bloqueado para nuevos lotes.   🟡 LISTO_RECOGIDA: avisado.   ⚪ ENTREGADO.   ⚠️ REVISAR: intervención humana.   🕒 TARDIO=RETENIDO: pagado pero retenido de producción hasta liberarlo (habla con el proveedor).')
    .setWrap(true).setVerticalAlignment('middle').setFontColor(COL.slate);
  sh.setRowHeight(32, 40);
}

function insertarGraficos(ss) {
  var sh = ss.getSheetByName(SH.DASH), data = ss.getSheetByName(SH.DATA);
  if (!sh || !data) return;
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });

  var dona = sh.newChart().asPieChart().setOption('pieHole', 0.55)
    .addRange(data.getRange('A1:B20')).setNumHeaders(1)
    .setOption('title', 'Distribución por estado').setOption('legend', { position: 'right' })
    .setOption('width', 430).setOption('height', 240).setPosition(11, 2, 0, 0).build();
  sh.insertChart(dona);

  var barras = sh.newChart().asColumnChart()
    .addRange(data.getRange('D1:E20')).setNumHeaders(1)
    .setOption('title', 'Camisetas conciliadas por talla').setOption('legend', { position: 'none' })
    .setOption('colors', [COL.orange]).setOption('width', 470).setOption('height', 240).setPosition(11, 6, 0, 0).build();
  sh.insertChart(barras);
}

function ordenarPestanas(ss) {
  var orden = [SH.HOWTO, SH.DASH, SH.CONFIG, SH.CATALOGO, SH.PEDIDOS, SH.LINEAS, SH.BANCO, SH.LOTES, SH.PROVEEDOR, SH.LOG];
  orden.forEach(function (n, i) { var s = ss.getSheetByName(n); if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); } });
  ss.setActiveSheet(ss.getSheetByName(SH.DASH));
}

/* ---- helpers de estilo ---- */

function hoja(ss, nombre, cabecera) {
  var sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  if (sh.getLastRow() === 0 && cabecera) sh.getRange(1, 1, 1, cabecera.length).setValues([cabecera]);
  return sh;
}
function banner(sh, a1, texto, size) {
  sh.getRange(a1).merge().setValue(texto).setBackground(COL.orange).setFontColor(COL.white)
    .setFontWeight('bold').setFontSize(size || 14).setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(sh.getRange(a1).getRow(), 40);
}
function sub(sh, a1, texto) { sh.getRange(a1).merge().setValue(texto).setFontColor(COL.slate).setFontStyle('italic'); }
function tarjeta(sh, fila, colIni, label, formula, tint, esMoneda) {
  var lab = sh.getRange(fila, colIni, 1, 2).merge();
  var val = sh.getRange(fila + 1, colIni, 1, 2).merge();
  lab.setValue(label).setBackground(tint).setFontColor(COL.slate).setFontSize(9).setFontWeight('bold')
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
  val.setFormula(formula).setBackground(tint).setFontColor(COL.ink).setFontSize(20).setFontWeight('bold')
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
  if (esMoneda) val.setNumberFormat(FMT_EUR);
  sh.setRowHeight(fila, 22); sh.setRowHeight(fila + 1, 40);
}

/* ===========================  MENÚ  =================================== */

function onOpen() {
  ui().createMenu('👕 Tienda Airbus 2026')
    .addItem('🛠️ Preparar backend (setup)', 'setupTiendaV4')
    .addItem('🎨 Reconstruir panel y formato', 'reconstruirEstetica')
    .addSeparator()
    .addItem('🏦 Conciliar banco', 'conciliarBanco')
    .addItem('✅ Confirmar PAGO del seleccionado (manual)', 'confirmarPagoSeleccion')
    .addItem('✅ Confirmar pagos por lista (CONFIRMAR_LOTE)', 'confirmarPagosPorLista')
    .addItem('📨 Avisar pedidos revisados (REVISADOS_LOTE)', 'avisarPedidosRevisadosPorLista')
    .addItem('⏳ Caducar pendientes vencidos', 'caducarPendientes')
    .addSeparator()
    .addItem('📦 Generar pedido a proveedor', 'generarPedidoProveedor')
    .addItem('🕒 Liberar pedidos tardíos (selección)', 'liberarPedidosTardiosSeleccion')
    .addItem('🧾 Refrescar RESUMEN_PROVEEDOR', 'refrescarResumenProveedor')
    .addItem('🔧 Reconstruir PROVEEDOR desde LINEAS', 'reconstruirProveedorDesdeLineas')
    .addItem('📥 Marcar lote recibido (seleccionado)', 'marcarLoteRecibidoSeleccion')
    .addItem('🤝 Marcar ENTREGADO (seleccionado)', 'marcarEntregadoSeleccion')
    .addSeparator()
    .addItem('📊 Actualizar panel', 'refrescarDashboard')
    .addItem('📗 Exportar a Excel (.xlsx)', 'exportarExcel')
    .addSeparator()
    .addItem('💾 Backup ahora (email fuera de Google)', 'backupDiario')
    .addItem('⏰ Programar backup diario', 'programarBackupDiario')
    .addSeparator()
    .addItem('✉️ Enviar emails de prueba', 'enviarEmailsPrueba')
    .addItem('📮 Configurar email (proveedor)', 'configurarEmailProveedor')
    .addItem('✉️ Reenviar confirmaciones fallidas', 'reenviarConfirmacionesFallidas')
    .addItem('✉️ Enviar confirmación a pagados (recovery)', 'enviarConfirmacionesPagoConciliado')
    .addItem('🔑 Mostrar TOKEN backend', 'mostrarToken')
    .addSeparator()
    .addItem('🧨 Resetear datos de prueba', 'resetearPruebas')
    .addToUi();
}

/* ===========================  UTILIDADES  ============================= */

function ui() { return SpreadsheetApp.getUi(); }
function limpiar(s, max) { return String(s == null ? '' : s).trim().slice(0, max || 200); }
function escapar(s) { return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function eur(n) { return (Number(n) || 0).toFixed(2).replace('.', ',') + ' €'; }
function parseImporte(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/\./g, '').replace(',', '.'); else s = s.replace(',', '.');
  return Number(s) || 0;
}
