# Conciliación bancaria — Tienda / Plataforma Solidaria (Huelga Airbus 2026)

Guía operativa del **backend** (Google Sheet + Apps Script) para cotejar las
transferencias con los pedidos. Complemento de `README.md` y `LANZAMIENTO.md`.

> Versión visual (tipo Power BI) del flujo completo conciliación → entrega:
> https://claude.ai/code/artifact/6070e092-8f48-4e26-a127-435097d9b412

---

## 1. Qué es y qué hace

La conciliación **casa cada ingreso del banco con su pedido** por dos criterios:
el **código `AIR26-XXXXX`** que el cliente pone en el concepto + el **importe exacto**.
Cuando cuadra: el pedido pasa a `PAGO_CONCILIADO` y se **envía el email de
confirmación** solo. Lo que no cuadra queda marcado para revisar a mano.

## 2. Archivos / hojas necesarias

Todo vive en el **Google Sheet** del backend (cuenta operativa), más `Code.gs`:

| Elemento | Para qué |
|---|---|
| Hoja **`MOVIMIENTOS_BANCO`** | Donde pegas el extracto. Rellenas tú: `FECHA · CONCEPTO · IMPORTE · REFERENCIA`. Rellena el script: `PEDIDO_DETECTADO · RESULTADO · PROCESADO · FECHA_CONCILIACION`. |
| Hoja **`PEDIDOS`** | Los pedidos (`ESTADO`, `TOTAL_EUR`, `CLIENT_REQUEST_ID`, `EMAIL`, `NOMBRE`…). Es contra lo que se casa. |
| Hoja **`CONFIG`** | `IBAN`, `BENEFICIARIO`, `PREFIJO` (=`AIR26`). |
| Hoja **`LOG`** | Registro de cada conciliación. |
| **`backend/Code.gs`** | Lógica: `conciliarBanco()`, `detectarId()`, `confirmarPagoSeleccion()`, `parseImporte()`. Se lanza desde el menú **👕 Tienda Airbus 2026**. |
| Export de tu **banco** | El extracto de ingresos del día. |

## 3. Cómo detecta el código

- `detectarId()` **aplana** el concepto (quita separadores) y busca `AIR26` + el
  número real. Tolera: `AIR26-00123`, `AIR26 00123`, `AIR26_00123`, `AIR2600123`,
  `AIR26.00123`, varios espacios, minúsculas y texto alrededor.
- **Bug corregido (sep-2026):** la versión previa (`normalizarId`) cogía el
  “26” de “AIR26” como número → conciliaba contra el pedido equivocado (o ninguno),
  incluso con el concepto correcto. Sustituida por `detectarId` (probada con 17 variantes).
- Para reducir errores, **el email y la web ya no exigen el guion**: si el banco no
  lo admite, vale sin él o con un espacio (lo importante es `AIR26` + el número).

## 4. Cómo casa el importe

- **Exacto al céntimo** (`parseImporte` admite `10,00`, `10.00` o `1.234,56`).
- **Incluye la aportación extra**: 10 € camiseta + 20 € extra = **30,00 €**. Si el
  ingreso no coincide → `REVISAR_IMPORTE`.
- Las **donaciones sin camiseta** casan igual (su total = la aportación).

## 5. Proceso diario — pasos

1. **Exporta** del banco los **ingresos** del día (solo abonos, no cargos).
2. **Pega los ingresos nuevos DEBAJO** de los anteriores en `MOVIMIENTOS_BANCO`
   (solo las 4 columnas tuyas; deja vacías las 4 de salida). **Se acumula, no se
   borra.** No re-pegues el extracto entero (duplicarías filas).
3. Menú → **🏦 Conciliar banco**.
4. Repasa la columna **`RESULTADO`**.
5. Resuelve a mano los `REVISAR_*` (sección 7).

## 6. Resultados posibles (`RESULTADO`)

| Resultado | Qué pasó | ¿Marca PROCESADO=SI? | Qué haces |
|---|---|---|---|
| `PAGO_CONCILIADO` | Casado + email enviado | **Sí** | Nada |
| `YA_PAGADO` | El pedido ya estaba pagado | **Sí** | Revisa si es ingreso duplicado |
| `REVISAR_IMPORTE (X vs Y)` | Código OK, importe distinto | No | Pagó de más/menos → decides |
| `REVISAR_SIN_CODIGO` | No se detectó “AIR26 + número” | No | Casar a mano por importe + nombre |
| `REVISAR_CODIGO_INEXISTENTE` | Código detectado que no existe | No | Buscar el pedido real y confirmar |
| `REVISAR_CADUCADO` | Pagó un pedido ya caducado | No | Decidir si se reactiva |

## 7. Confirmar a mano (para cualquier `REVISAR_*`)

1. En `PEDIDOS`, localiza el pedido por **importe (`TOTAL_EUR`) + nombre/email**
   que aparezca en el concepto o la referencia.
2. **Selecciona su fila** → menú → **✅ Confirmar PAGO del seleccionado**. Pasa a
   `PAGO_CONCILIADO` y manda el email.
3. En `MOVIMIENTOS_BANCO`, pon **`SI`** en `PROCESADO` de ese movimiento.

## 8. Reglas clave

- **Acumular, no borrar.** El script solo procesa filas con `PROCESADO` vacío; las
  `SI` se saltan → re-conciliar es **seguro** e idempotente.
- **Pega solo lo nuevo** cada día.
- **Los `REVISAR_*` reaparecen** hasta que los cierres con `PROCESADO=SI`: son tu
  pila de pendientes.
- **No borres** las 4 columnas de salida de `MOVIMIENTOS_BANCO`.
- **Solo ingresos** (los cargos negativos no casan).

## 9. Verificación realizada

Probado en producción con dos conceptos "malos" a propósito — `AIR2600601`
(sin guion) y `AIR26_00565` (guion bajo) — y **ambos conciliaron** correctamente.
Los dos pedidos de prueba se revirtieron luego a `CADUCADO`.

## 10. Pedidos tardíos (retenidos de producción)

Desde el **domingo 6-sep 21:00** (fecha límite para llegar a la marcha del 12-S),
todo pedido nuevo se marca **tardío**. Un tardío **se cobra y se concilia igual**
que cualquiera —no cambia nada del pago—, pero queda **retenido**: **no entra solo
al proveedor**. Antes de producirlos hay que **hablar con el proveedor** para ver
si llegan a tiempo.

Cómo funciona, en la práctica:

- **La marca es una columna nueva en `PEDIDOS`: `TARDIO`.** Valores:
  `RETENIDO` (rojo, retenido), `LIBERADO` (verde, ya se produce) o vacío (a tiempo).
  Lo pone el **backend** al crear el pedido, comparando con `CONFIG → AVISO_FECHA_LIMITE`
  (no decide el navegador). Esa fecha debe **coincidir** con la de `config.js` de la web.
- **Conciliar** un tardío es normal: pasa a `PAGO_CONCILIADO` y recibe su email. No hay
  que hacer nada distinto en el proceso diario de la sección 5.
- **📦 Generar pedido a proveedor** **salta** los `RETENIDO` y te dice cuántos ha
  dejado fuera. Los pedidos a tiempo se producen; los tardíos esperan.
- **Cuando el proveedor confirme plazos:** en `PEDIDOS`, **selecciona las filas** de
  los pedidos tardíos (una o varias) → menú → **🕒 Liberar pedidos tardíos (selección)**.
  Pasan a `LIBERADO` y entran en el **siguiente** *Generar pedido a proveedor*.
  Liberar **no** cambia el estado ni envía emails: solo levanta la retención.

## 11. Desplegar un cambio del backend

`conciliarBanco` vive en `Code.gs`. Para cualquier cambio: pega el `Code.gs` →
*Implementar → Gestionar implementaciones → ✏️ Editar → Nueva versión → Implementar*.
La conciliación es **manual**, así que tocarla **no afecta** al alta de pedidos.
Rollback: en *Gestionar implementaciones* vuelve a la versión anterior.

> **Al desplegar la columna `TARDIO` por primera vez:** tras pegar el `Code.gs`,
> ejecuta una vez el menú **👕 → 🎨 Reconstruir panel y formato**. Añade la cabecera
> `TARDIO` a `PEDIDOS` y sus colores, sin tocar los datos. (El *setup* no reescribe la
> cabecera si la hoja ya tiene pedidos; *Reconstruir panel y formato* sí.) Los pedidos
> anteriores quedan con `TARDIO` vacío = a tiempo, que es lo correcto.

---

## Después de conciliar (mapa, fuera de esta guía)

`PAGO_CONCILIADO` → **📦 Generar pedido a proveedor** → `ENVIADO_PROVEEDOR` →
**📥 Marcar lote recibido** (hoja `LOTES`) → `LISTO_RECOGIDA` (+ email) →
**🤝 Marcar ENTREGADO** (hoja `PEDIDOS`) → `ENTREGADO`.

Los pedidos **`TARDIO = RETENIDO`** se saltan el primer paso hasta liberarlos
(**🕒 Liberar pedidos tardíos**); ver sección 10.
