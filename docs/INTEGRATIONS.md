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
