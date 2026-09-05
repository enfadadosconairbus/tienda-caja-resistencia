/* =============================================================================
   app.js — lógica de la tienda (carrito, validación, envío)
   Cableado contra los IDs/clases de index.html. Sin dependencias externas.
   ========================================================================== */
(function () {
  'use strict';

  var CFG = window.TIENDA_CONFIG || {};
  var CATALOG = (CFG.CATALOG || []).filter(function (p) { return p.activo !== false; });
  var MAX_UNITS = CFG.MAX_UNITS || 20;

  // ---- estado ----------------------------------------------------------------
  var state = {
    selectedSku: CATALOG.length ? CATALOG[0].sku : null,
    qty: 1,
    donation: 0,
    lines: {}          // sku -> { sku, talla, precio, qty }
  };

  // client_request_id estable para reintentos (idempotencia en el backend)
  var CRID = newRequestId();

  // ---- helpers de DOM --------------------------------------------------------
  function $(id) { return document.getElementById(id); }
  function eur(n) { return (Number(n) || 0).toFixed(2).replace('.', ',') + ' €'; }
  function bySku(sku) {
    for (var i = 0; i < CATALOG.length; i++) if (CATALOG[i].sku === sku) return CATALOG[i];
    return null;
  }
  function newRequestId() {
    try { if (window.crypto && crypto.randomUUID) return crypto.randomUUID(); } catch (e) {}
    return 'crid-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }

  // ---- render: selector de talla --------------------------------------------
  function renderSizes() {
    var chips = $('sizeChips');
    var select = $('sizeSelect');
    if (!chips || !select) return;
    chips.innerHTML = '';
    select.innerHTML = '';
    CATALOG.forEach(function (p) {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'size-chip' + (p.sku === state.selectedSku ? ' active' : '');
      b.textContent = p.talla;
      b.setAttribute('aria-pressed', p.sku === state.selectedSku ? 'true' : 'false');
      b.addEventListener('click', function () { selectSize(p.sku); });
      chips.appendChild(b);

      var o = document.createElement('option');
      o.value = p.sku; o.textContent = p.talla;
      if (p.sku === state.selectedSku) o.selected = true;
      select.appendChild(o);
    });
    select.onchange = function () { selectSize(select.value); };
  }

  function selectSize(sku) {
    state.selectedSku = sku;
    renderSizes();
  }

  // ---- cantidad --------------------------------------------------------------
  function setQty(n) {
    state.qty = Math.max(1, Math.min(MAX_UNITS, n));
    var out = $('qtyOutput');
    if (out) out.textContent = String(state.qty);
  }

  // ---- carrito ---------------------------------------------------------------
  function unitsInCart() {
    var t = 0; for (var k in state.lines) t += state.lines[k].qty; return t;
  }
  function addToCart() {
    var p = bySku(state.selectedSku);
    if (!p) return;
    var free = MAX_UNITS - unitsInCart();
    if (free <= 0) { flashCatalog('Has alcanzado el máximo de ' + MAX_UNITS + ' unidades por pedido.'); return; }
    var add = Math.min(state.qty, free);
    var line = state.lines[p.sku] || { sku: p.sku, talla: p.talla, precio: p.precio, qty: 0 };
    line.qty += add;
    state.lines[p.sku] = line;
    setQty(1);
    renderCart();
    flashCatalog(add < state.qty ? 'Añadidas ' + add + ' uds. (tope ' + MAX_UNITS + ').' : 'Añadido al pedido.');
  }
  function removeLine(sku) { delete state.lines[sku]; renderCart(); }
  function changeLineQty(sku, delta) {
    var line = state.lines[sku]; if (!line) return;
    var others = unitsInCart() - line.qty;
    line.qty = Math.max(0, Math.min(MAX_UNITS - others, line.qty + delta));
    if (line.qty === 0) delete state.lines[sku];
    renderCart();
  }

  function renderCart() {
    var wrap = $('cartLines');
    var empty = $('cartEmpty');
    var skus = Object.keys(state.lines);
    if (wrap) wrap.innerHTML = '';
    if (empty) empty.hidden = skus.length > 0;

    var productsTotal = 0;
    skus.forEach(function (sku) {
      var l = state.lines[sku];
      var sub = l.precio * l.qty;
      productsTotal += sub;
      if (!wrap) return;
      var row = document.createElement('div');
      row.className = 'cart-line';
      row.innerHTML =
        '<div class="cl-main"><strong>' + CFG.PRODUCT.nombre + ' · ' + l.talla + '</strong>' +
        '<span class="cl-sub">' + eur(l.precio) + ' / ud.</span></div>' +
        '<div class="cl-qty">' +
          '<button type="button" class="lineMinus" aria-label="Restar">−</button>' +
          '<span>' + l.qty + '</span>' +
          '<button type="button" class="linePlus" aria-label="Sumar">+</button>' +
        '</div>' +
        '<div class="cl-sum">' + eur(sub) + '</div>' +
        '<button type="button" class="cl-del" aria-label="Quitar">✕</button>';
      row.querySelector('.lineMinus').addEventListener('click', function () { changeLineQty(sku, -1); });
      row.querySelector('.linePlus').addEventListener('click', function () { changeLineQty(sku, 1); });
      row.querySelector('.cl-del').addEventListener('click', function () { removeLine(sku); });
      wrap.appendChild(row);
    });

    var donation = Number(state.donation) || 0;
    var grand = productsTotal + donation;
    setText('productsTotal', eur(productsTotal));
    setText('donationTotal', eur(donation));
    setText('grandTotal', eur(grand));

    var units = unitsInCart();
    setText('unitCount', units + (units === 1 ? ' ud.' : ' uds.'));

    // barra móvil
    var sticky = $('stickyCart');
    if (sticky) {
      sticky.hidden = units === 0;
      setText('stickyUnits', units + (units === 1 ? ' ud.' : ' uds.'));
      setText('stickyTotal', eur(grand));
    }
  }

  function setText(id, txt) { var el = $(id); if (el) el.textContent = txt; }
  function flashCatalog(msg) { var el = $('catalogStatus'); if (el) el.textContent = msg; }

  // ---- aportación ------------------------------------------------------------
  function wireDonations() {
    var box = $('donationButtons'); if (!box) return;
    var custom = $('donationCustom');
    function clearButtons() { Array.prototype.forEach.call(box.querySelectorAll('button'), function (b) { b.classList.remove('active'); }); }
    box.addEventListener('click', function (e) {
      var btn = e.target.closest('button[data-donation]'); if (!btn) return;
      state.donation = Number(btn.getAttribute('data-donation')) || 0;
      if (custom) custom.value = '';               // el botón fijo manda: limpia el campo libre
      clearButtons();
      btn.classList.add('active');
      renderCart();
    });
    if (custom) custom.addEventListener('input', function () {
      var n = parseFloat(String(custom.value).replace(',', '.')) || 0;   // admite coma o punto
      state.donation = Math.max(0, Math.round(n * 100) / 100);           // conserva céntimos
      clearButtons();
      renderCart();
    });
  }

  // ---- envío -----------------------------------------------------------------
  function validate(form) {
    if (form.website && form.website.value) return '__bot__'; // honeypot
    var units = unitsInCart();
    var donation = Number(state.donation) || 0;
    if (units === 0 && donation <= 0) return 'Añade una camiseta o indica una aportación.';
    var nombre = form.nombre.value.trim();
    var apellidos = form.apellidos.value.trim();
    var email = form.email.value.trim();
    var telefono = form.telefono.value.trim();
    if (!nombre || !apellidos) return 'Indica tu nombre y apellidos.';
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return 'Revisa tu email.';
    if (telefono.replace(/\D/g, '').length < 9) return 'Revisa tu teléfono.';
    if (form.site && !form.site.value) return 'Elige tu site de recogida.';
    if (!form.privacy.checked) return 'Debes aceptar la información de privacidad.';
    return null;
  }

  function buildPayload(form) {
    return {
      action: 'crear_pedido',
      client_request_id: CRID,
      cliente: {
        nombre: form.nombre.value.trim(),
        apellidos: form.apellidos.value.trim(),
        email: form.email.value.trim(),
        telefono: form.telefono.value.trim()
      },
      site: form.site ? form.site.value : '',
      lineas: Object.keys(state.lines).map(function (sku) {
        var l = state.lines[sku];
        return { producto: CFG.PRODUCT.nombre, sku: l.sku, talla: l.talla, cantidad: l.qty };
      }),
      aportacion: Number(state.donation) || 0,
      recogida: CFG.PICKUP || ''
    };
  }

  function onSubmit(e) {
    e.preventDefault();
    var form = e.target;
    var err = $('formError');
    if (CFG.APORTACIONES_ACTIVAS === false) {
      showError(err, 'Las aportaciones aún no están activas. Se abrirán muy pronto, cuando la cuenta bancaria esté disponible.');
      return;
    }
    var problem = validate(form);
    if (problem === '__bot__') return;          // silencio ante bots
    if (problem) { showError(err, problem); return; }
    hideError(err);

    var btn = $('submitOrder');
    setBusy(btn, true);

    var payload = buildPayload(form);

    if (CFG.DEMO_MODE) {
      setTimeout(function () {
        setBusy(btn, false);
        showSuccess(demoResponse(payload));
      }, 500);
      return;
    }

    sendOrder(payload).then(function (resp) {
      setBusy(btn, false);
      if (resp && resp.ok) showSuccess(resp);
      else showError(err, (resp && resp.error) || 'No se pudo registrar el pedido. Inténtalo de nuevo.');
    }).catch(function () {
      setBusy(btn, false);
      showError(err, 'Fallo de conexión. Revisa tu conexión e inténtalo de nuevo.');
    });
  }

  function sendOrder(payload) {
    return fetch(CFG.API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // evita preflight innecesario
      body: JSON.stringify(payload)
    }).then(function (r) { return r.json(); });
  }

  function demoResponse(payload) {
    var productos = payload.lineas.reduce(function (s, l) {
      var p = bySku(l.sku); return s + (p ? p.precio : 0) * l.cantidad;
    }, 0);
    return {
      ok: true,
      order_id: 'AIR26-00042',
      total: productos + (payload.aportacion || 0),
      beneficiario: CFG.DEMO_BENEFICIARIO,
      iban: CFG.DEMO_IBAN,
      concepto: 'AIR26-00042',
      demo: true
    };
  }

  // ---- pantalla de éxito -----------------------------------------------------
  function showSuccess(resp) {
    setText('successOrderId', resp.order_id || '');
    setText('successBeneficiary', resp.beneficiario || '');
    setText('successIban', resp.iban || '');
    setText('successTotal', eur(resp.total));
    setText('successConcept', resp.concepto || resp.order_id || '');
    // Site de recogida: Getafe e Illescas se recogen en Getafe; el resto en su site.
    var sel = document.getElementById('siteSelect');
    var siteSel = sel ? String(sel.value || '').trim() : '';
    var pickupSite = /^(getafe|illescas)$/i.test(siteSel) ? 'Getafe' : (siteSel || 'tu site');
    setText('successPickup', pickupSite);
    var ok = $('success');
    var pedido = $('pedido');
    if (pedido) pedido.hidden = true;
    if (ok) { ok.hidden = false; ok.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
    var sticky = $('stickyCart'); if (sticky) sticky.hidden = true;
  }

  // ---- copiar ----------------------------------------------------------------
  function wireCopy() {
    document.addEventListener('click', function (e) {
      var btn = e.target.closest('.copy-btn'); if (!btn) return;
      var target = $(btn.getAttribute('data-copy')); if (!target) return;
      var text = target.textContent.trim();
      var done = function () { var old = btn.textContent; btn.textContent = 'COPIADO'; setTimeout(function () { btn.textContent = old; }, 1500); };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(done).catch(function () { legacyCopy(text); done(); });
      } else { legacyCopy(text); done(); }
    });
  }
  function legacyCopy(text) {
    var ta = document.createElement('textarea'); ta.value = text; ta.setAttribute('readonly', '');
    ta.style.position = 'absolute'; ta.style.left = '-9999px'; document.body.appendChild(ta);
    ta.select(); try { document.execCommand('copy'); } catch (e) {} document.body.removeChild(ta);
  }

  // ---- textos legales (modales) ---------------------------------------------
  function wireLegal() {
    var openers = document.querySelectorAll('[data-legal]');
    Array.prototype.forEach.call(openers, function (btn) {
      btn.addEventListener('click', function () {
        var dlg = $('dlg-' + btn.getAttribute('data-legal'));
        if (!dlg) return;
        if (dlg.showModal) dlg.showModal(); else dlg.setAttribute('open', '');
      });
    });
    var dialogs = document.querySelectorAll('dialog.legal-dialog');
    Array.prototype.forEach.call(dialogs, function (dlg) {
      Array.prototype.forEach.call(dlg.querySelectorAll('[data-close]'), function (b) {
        b.addEventListener('click', function () { dlg.close ? dlg.close() : dlg.removeAttribute('open'); });
      });
      // cerrar al pulsar el fondo (fuera del contenido)
      dlg.addEventListener('click', function (e) { if (e.target === dlg) dlg.close && dlg.close(); });
    });
  }

  // ---- utilidades UI ---------------------------------------------------------
  function showError(el, msg) { if (!el) return; el.textContent = msg; el.hidden = false; el.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
  function hideError(el) { if (!el) return; el.hidden = true; el.textContent = ''; }
  function setBusy(btn, busy) {
    if (!btn) return;
    btn.disabled = busy;
    if (busy) { btn.dataset.label = btn.textContent; btn.textContent = 'ENVIANDO…'; }
    else if (btn.dataset.label) { btn.textContent = btn.dataset.label; }
  }

  // ---- avisos de entrega (fecha / stock) -------------------------------------
  function evaluarAvisoEntrega(tardePorStock) {
    var limite = CFG.AVISO_FECHA_LIMITE ? new Date(CFG.AVISO_FECHA_LIMITE) : null;
    var tardePorFecha = !!(limite && !isNaN(limite.getTime()) && new Date() >= limite);
    var tarde = tardePorFecha || !!tardePorStock;
    var html, urgente;
    if (tarde) {
      urgente = true;
      html = '<span aria-hidden="true">⚠️</span> <strong>Posible entrega después de la marcha.</strong> Por los plazos de producción, tu camiseta podría llegar <strong>después de la marcha al Ministerio del 12-Septiembre</strong>. Tu aportación sigue sosteniendo la caja de resistencia.';
    } else {
      urgente = false;
      html = '<span aria-hidden="true">⏱️</span> <strong>Pide antes del domingo 6 de septiembre a las 21:00</strong> para recibir la camiseta antes de la <strong>marcha al Ministerio del 12-Septiembre</strong>. Después de esa hora podría llegarte tras la marcha (producción 4-5 días).';
    }
    [$('avisoEntregaPedido'), $('avisoEntregaExito')].forEach(function (el) {
      if (!el) return;
      el.innerHTML = html;
      if (urgente) el.classList.add('urgente'); else el.classList.remove('urgente');
      el.hidden = false;
    });
  }

  function comprobarStock() {
    if (CFG.APORTACIONES_ACTIVAS !== true) return;        // solo con la adquisición abierta
    var umbral = Number(CFG.AVISO_STOCK_UMBRAL) || 0; if (umbral <= 0) return;
    fetch(CFG.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'estado' }) })
      .then(function (r) { return r.json(); })
      .then(function (resp) { if (resp && resp.ok && Number(resp.camisetas) >= umbral) evaluarAvisoEntrega(true); })
      .catch(function () {});
  }

  // ---- init ------------------------------------------------------------------
  function init() {
    // textos de config
    var contact = $('contactEmail');
    if (contact && CFG.CONTACT_EMAIL) { contact.textContent = CFG.CONTACT_EMAIL; contact.href = 'mailto:' + CFG.CONTACT_EMAIL; }
    var pickup = $('pickupText'); if (pickup && CFG.PICKUP) pickup.textContent = CFG.PICKUP;

    renderSizes();
    setQty(1);
    renderCart();
    wireDonations();
    wireCopy();
    wireLegal();

    var minus = $('minusQty'); if (minus) minus.addEventListener('click', function () { setQty(state.qty - 1); });
    var plus = $('plusQty'); if (plus) plus.addEventListener('click', function () { setQty(state.qty + 1); });
    var add = $('addToCart'); if (add) add.addEventListener('click', addToCart);
    var form = $('checkoutForm'); if (form) form.addEventListener('submit', onSubmit);
    var go = $('stickyGo'); if (go) go.addEventListener('click', function () {
      var p = $('pedido'); if (p) p.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });

    // Pre-lanzamiento: aportaciones aún no activas (falta la cuenta bancaria).
    if (CFG.APORTACIONES_ACTIVAS === false) {
      var av = $('avisoBanner'); if (av) av.hidden = false;
      var oc = $('ordenNoDisponible'); if (oc) oc.hidden = false;
      var sub = $('submitOrder'); if (sub) { sub.disabled = true; sub.textContent = 'DISPONIBLE MUY PRONTO'; }
    }

    evaluarAvisoEntrega(false);   // aviso de entrega por fecha (visible desde ya)
    comprobarStock();             // y por stock, si la adquisición está abierta

    if (CFG.DEMO_MODE) flashCatalog('MODO DEMO: los pedidos no se registran. Pon DEMO_MODE:false para producción.');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
