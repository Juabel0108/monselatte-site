# Panel administrativo — primera versión

## Objetivo

Permitir que Monselatte prepare, revise y envíe una cotización desde una interfaz privada sin editar directamente las columnas de Google Sheets.

## Arquitectura recomendada

El panel no debe publicarse en GitHub Pages. Debe vivir en un proyecto o despliegue de Google Apps Script separado, configurado para que solo la cuenta de Monselatte tenga acceso. La hoja `Leads` continúa siendo la fuente de datos durante esta fase.

El Apps Script público que recibe el formulario permanece sin cambios. El panel privado reutiliza la misma operación de generación y envío; no crea una segunda implementación de esa lógica.

## Implementación privada actual

- Descripción: `Panel administrativo privado · confirmación y envío seguro`.
- Versión: 36.
- Acceso: `Only myself`.
- Ejecuta como: `monselattepr@gmail.com`.
- URL: `https://script.google.com/macros/s/AKfycbxP1nx7ArXxunOVGvpBFrKM29NOC4DuBPtDSB8PKvNTaAlhMQ0YmdLqP_HMAWm3A8EOGA/exec`

La implementación pública del formulario continúa separada en la versión 31 y conserva su URL anterior.

## Primera pantalla

- Bandeja de solicitudes pendientes.
- Búsqueda por nombre, email o número de cotización.
- Estado visible: nueva, preparada, enviada o con error.
- Fecha del evento, municipio, invitados y menú solicitado.

## Preparación de la cotización

- Precio de la barra.
- Cantidad y precio de bebidas frías.
- Cantidad y precio de matcha.
- Cantidad y precio de espresso martini.
- Cantidad y precio de carajillo.
- Calculadora visual de precios.
- Total oficial introducido manualmente.
- Vista previa antes de enviar.

## Acción de envío

El botón debe decir `Generar PDF y enviar`. Antes de ejecutar:

1. Mostrar destinatario y total oficial.
2. Pedir confirmación explícita.
3. Deshabilitar dobles clics mientras procesa.
4. Mostrar el resultado real del servidor.
5. Escribir `pdfUrl` y `sent_at` únicamente después de `{ sent: true }`.
6. Mostrar el error de Gmail o del adjunto sin convertirlo en éxito.

## Primera entrega segura

1. Crear el proyecto administrativo privado.
2. Conectarlo inicialmente a una hoja o fila de prueba.
3. Implementar lectura y edición sin envío. ✅
4. Añadir vista previa del PDF. ✅
5. Activar el envío solamente después de una prueba controlada. ✅ Probado con `ML-2609-TEST`: Gmail aceptó el envío, se generó un PDF de una página y `sent_at` se registró después del éxito explícito.
6. Mantener el trigger actual como respaldo hasta comprobar el panel.
