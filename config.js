/* =============================================================================
   config.js — CONFIGURACIÓN PÚBLICA de la tienda
   -----------------------------------------------------------------------------
   Este fichero SÍ se publica en GitHub. Por eso aquí NO va ningún secreto:
   ni el token del backend, ni la URL /exec de Apps Script, ni el IBAN real.

   - El IBAN, el beneficiario y el importe los devuelve el BACKEND al crear el
     pedido (así el navegador nunca decide el dinero).
   - El catálogo de aquí es solo para PINTAR tallas y precios en pantalla; el
     backend vuelve a validar cada precio contra la hoja CATALOGO (fuente de
     verdad). Si alguien manipula el JavaScript, el servidor manda.

   ANTES DE PUBLICAR:
     1) API_URL  -> pon la URL real de tu Cloudflare Worker.
     2) DEMO_MODE -> false.
   ========================================================================== */
window.TIENDA_CONFIG = {
  // URL pública del Cloudflare Worker (el proxy). Nunca la URL /exec directa.
  API_URL: 'https://caja-resistencia.enfadadosconairbus-contacto.workers.dev',

  // true  = no llama al backend; simula el pedido en el navegador (para probar diseño).
  // false = pedidos reales contra el Worker. Ponlo en false antes de compartir la URL.
  DEMO_MODE: false,

  // Aportaciones DESACTIVADAS hasta tener la cuenta bancaria lista.
  // Cuando la tengáis: pon true, guarda y haz push → se abre todo (y desaparece el aviso).
  APORTACIONES_ACTIVAS: true,

  // Avisos de entrega ("puede llegar después de la marcha del 12-S"):
  // 1) por fecha: pedidos a partir de este instante (domingo 6-sep, 21:00 hora peninsular).
  //    OJO: este valor es solo para el AVISO de la web. Quien decide de verdad si un
  //    pedido es "tardío" (y lo RETIENE de producción) es el backend, con su propia
  //    copia en CONFIG → AVISO_FECHA_LIMITE. Mantén las dos fechas iguales.
  AVISO_FECHA_LIMITE: '2026-09-06T21:00:00+02:00',
  // 2) por stock: cuando se alcanzan estas camisetas pedidas (solo si las aportaciones están activas).
  AVISO_STOCK_UMBRAL: 2500,

  // Producto único de momento (arquitectura preparada para multiproducto).
  PRODUCT: { nombre: 'Camiseta', skuPrefix: 'CAMISETA' },

  MAX_UNITS: 20,               // tope de unidades por pedido
  DONATIONS: [0, 5, 10, 20, 50],

  // Catálogo visible (talla, medidas, precio). El backend re-valida el precio.
  CATALOG: [
    { sku: 'CAMISETA-XS',  talla: 'XS',  ancho: 46, alto: 66, precio: 10, activo: true },
    { sku: 'CAMISETA-S',   talla: 'S',   ancho: 49, alto: 69, precio: 10, activo: true },
    { sku: 'CAMISETA-M',   talla: 'M',   ancho: 52, alto: 71, precio: 10, activo: true },
    { sku: 'CAMISETA-L',   talla: 'L',   ancho: 55, alto: 73, precio: 10, activo: true },
    { sku: 'CAMISETA-XL',  talla: 'XL',  ancho: 58, alto: 75, precio: 10, activo: true },
    { sku: 'CAMISETA-2XL', talla: '2XL', ancho: 62, alto: 77, precio: 10, activo: true },
    { sku: 'CAMISETA-3XL', talla: '3XL', ancho: 66, alto: 79, precio: 10, activo: true },
    { sku: 'CAMISETA-4XL', talla: '4XL', ancho: 70, alto: 81, precio: 10, activo: true },
    { sku: 'CAMISETA-5XL', talla: '5XL', ancho: 74, alto: 83, precio: 10, activo: true }
  ],

  PICKUP: 'Getafe - Factoría Airbus - Puerta Sur / Puerta Norte (Asamblea de trabajadores en Huelga)',
  CONTACT_EMAIL: 'enfadadosconairbus.contacto@gmail.com',

  // Solo se usan cuando DEMO_MODE = true, para poder ver la pantalla final.
  // En pedidos reales estos datos llegan del backend (hoja CONFIG).
  DEMO_BENEFICIARIO: 'Caja de Resistencia Huelga Airbus 2026 - Sindicato Útil',
  DEMO_IBAN: 'ESXX XXXX XXXX XXXX XXXX XXXX'
};
