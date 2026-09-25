# Integración de cotizaciones

## Contrato del formulario público

El formulario de `index.html` envía un `POST` al Web App de Google Apps Script mediante `fetch`, con `mode: no-cors` y un cuerpo JSON. La URL se obtiene del meta tag `sheet-webapp-url`; no debe cambiarse sin actualizar y volver a probar el despliegue.

`main.js` construye el payload desde `FormData`, conserva los nombres de los campos y combina los valores múltiples de menú en texto separado por comas. También incorpora los parámetros UTM. El Apps Script recibe el JSON en `doPost`, valida la solicitud, añade una fila en `Leads` y envía los correos iniciales.

No cambiar sin una prueba de regresión completa:

- URL o método del endpoint.
- `mode: no-cors` o formato JSON del cuerpo.
- atributos `name` e IDs usados por `main.js`.
- nombres o combinación de campos del payload.
- mensajes, redirecciones o comportamiento posterior al envío.

## Flujo administrativo actual

1. La solicitud queda registrada en la hoja `Leads`.
2. El administrador escribe los precios y cantidades.
3. `total_calculado` suma visualmente los precios, pero no modifica `precio`.
4. El administrador escribe manualmente `precio`, que es el total oficial.
5. Al marcar `aprobado = si`, el trigger instalable `onSheetEdit` inicia el flujo.
6. `onLeadChange` asigna el número si falta, genera el PDF, envía el correo y escribe `sent_at` solo tras éxito explícito.
7. Al editar `deposito` en una cotización aprobada con PDF, se genera y envía la actualización.
8. El panel permite actualizar manualmente `estado` sin enviar correo; un envío exitoso lo establece automáticamente en `ENVIADO`.

Las nuevas solicitudes se escriben en la primera fila sin identidad de cliente (`timestamp`, `nombre` y `email`). Esto evita que fórmulas extendidas en columnas administrativas obliguen a `appendRow` a saltar cientos de filas.

## Google Apps Script

- Proyecto: `Monselatte Cotizaciones`.
- Hoja: `Leads`.
- Trigger requerido: `onSheetEdit`, origen Spreadsheet, evento On edit.
- El código fuente de referencia es `apps-script.gs`; el código ejecutado vive en el editor de Google Apps Script.
- Propiedades requeridas: `SPREADSHEET_ID`, `LOGO_ID`, `QUOTES_FOLDER_ID`, `TERMS_PDF_ID` y `STRIPE_DEPOSIT_URL`.

## Reglas de seguridad del envío

- `sendQuoteEmail` falla si el destinatario está vacío.
- Los fallos de Gmail o del PDF de términos se registran y se relanzan.
- Un envío solo es exitoso cuando devuelve `{ sent: true }`.
- `sent_at` solo puede escribirse después de esa confirmación explícita.
- El panel administrativo futuro debe reutilizar esta misma operación; no debe implementar un segundo flujo de envío.

## Columnas administrativas de precios

- `precio_barra`
- `cantidad_bebidas_frias`, `precio_bebidas_frias`
- `cantidad_matcha`, `precio_matcha`
- `cantidad_espresso_martini`, `precio_espresso_martini`
- `cantidad_carajillo`, `precio_carajillo`
- `total_calculado` (solo referencia visual)
- `precio` (total oficial introducido manualmente)

## Estado del envío público con `no-cors`

El formulario valida antes de iniciar el POST, bloquea intentos simultáneos y
muestra «Enviando solicitud…» con `aria-busy` durante la promesa de `fetch`.
Solo después de su resolución muestra el panel de éxito existente. Si la
promesa rechaza por un error de red, restaura el botón y el foco, conserva los
datos y permite reintentar; muestra un aviso en #errors (aria-live polite) y
conserva el registro del error en consola.

**Una respuesta opaca de `mode: no-cors` no confirma que Apps Script haya
validado o guardado la solicitud ni enviado emails.** No permite leer el cuerpo
ni verificar el estado HTTP. El panel indica «SOLICITUD ENVIADA» para procesamiento y condiciona la
confirmación por email al procesamiento correcto. No afirma recepción
confirmada por el backend. No se ha cambiado el endpoint ni la arquitectura.

`generate_lead` y `contact` registran un intento validado cuyo `fetch` ya se
inició; no prueban procesamiento remoto. Un reintento tras rechazo de red es un
nuevo intento. No se disparan por errores de validación ni por dobles submits
bloqueados. Los nombres, parámetros y la serialización del payload se conservan.

Añasco falta tanto en el select como en `PUERTO_RICO_MUNICIPIOS` del código local
de Apps Script. No se añade al frontend mientras ese contrato no lo permita.
Se mantienen el control 2–8 horas, la validación existente 1–12 y la exclusión
de intervalos que cruzan medianoche; esta fase no reconcilia esos rangos.


## Texto externo al escribir en Sheets

`safeSheetText_` protege centralmente los valores de `appendRowByHeader_`
(Leads, incluidos menú y atribución) y las filas de `logSpam`.
Ante texto que comienza con =, +, - o @, incluso precedido de whitespace
o controles iniciales, antepone el marcador de texto de Sheets (apóstrofo).
No recorta ni sustituye el contenido y conserva los tipos no textuales.
La normalización previa de `sanitizeInput` no cambia.

Las otras escrituras usan encabezados/estados controlados, cantidades
convertidas a Number y validadas, fechas Date, números de cotización
generados internamente y URLs generadas por Drive.

Pruebas locales: `node tests/sheets-text-safety.test.js`,
`npm run test:quote` y `tests/wizard-hardening.test.js` con Playwright
y transporte simulado. Ninguna requiere escribir en la hoja real.
Este cambio local requiere una actualización posterior autorizada del código
ejecutado en Google Apps Script; no publica ni modifica el despliegue actual.
