AGENTS.md

Objetivo del proyecto

Monselatte es una landing page comercial para presentar el servicio de coffee bar móvil, generar confianza y recibir solicitudes de cotización.

El sitio debe optimizarse para:

1. Explicar claramente el servicio.
2. Mostrar eventos y fotografías reales.
3. Generar solicitudes de cotización.
4. Facilitar contacto por WhatsApp y teléfono.
5. Mantener buena experiencia móvil.
6. Mantener intacta la automatización existente con Google Apps Script.

Regla crítica de integración

El formulario de cotización está conectado a Google Apps Script y Google Sheets.

No modificar sin autorización explícita:

* La URL del endpoint.
* El método de envío.
* Los nombres de los campos.
* Los atributos name.
* Los IDs utilizados por JavaScript.
* La estructura del payload.
* Los nombres de parámetros.
* La lógica de cálculo.
* Las funciones que ejecutan fetch.
* Los mensajes de éxito o error.
* Las redirecciones posteriores al envío.
* La integración con Google Sheets.

Los estilos visuales del formulario pueden modificarse siempre que no cambie su contrato funcional.

Antes de editar cualquier archivo relacionado con el formulario:

1. Lee docs/INTEGRATIONS.md.
2. Identifica todos los campos y parámetros.
3. Documenta el comportamiento actual.
4. Crea o ejecuta pruebas de regresión.
5. Conserva la compatibilidad con Apps Script.

Después de editar:

1. Comprueba que todos los campos siguen existiendo.
2. Comprueba que los atributos name no cambiaron.
3. Comprueba que el payload mantiene las mismas claves.
4. Ejecuta la prueba de envío.
5. Confirma que se muestra el estado de éxito o error correcto.

Flujo de trabajo

Antes de realizar cambios:

1. Inspecciona el repositorio completo.
2. Lee los archivos dentro de docs/.
3. Ejecuta la página actual.
4. Identifica qué archivos controlan la sección solicitada.
5. Revisa el estado de Git.
6. No modifiques archivos no relacionados.

Diseño

La página debe sentirse:

* Elegante.
* Editorial.
* Cálida.
* Profesional.
* Relacionada con café y eventos.
* Diseñada específicamente para Monselatte.
* Visualmente limpia, pero no vacía.

Priorizar:

* Fotografías reales.
* Jerarquía tipográfica.
* Espacios amplios.
* Contenido fácil de escanear.
* CTA de cotización visible.
* Experiencia móvil.
* Prueba social.
* Información real del negocio.

Evitar:

* Apariencia de startup tecnológica.
* Gradientes genéricos.
* Glassmorphism.
* Exceso de tarjetas.
* Sombras fuertes.
* Bordes completamente redondeados en todo.
* Íconos decorativos sin función.
* Animaciones excesivas.
* Texto promocional genérico.
* Cambios de marca sin justificación.

Responsive

Verificar como mínimo:

* 375 × 812.
* 768 × 1024.
* 1024 × 768.
* 1440 × 900.

No debe existir:

* Overflow horizontal.
* Texto cortado.
* Imágenes deformadas.
* Botones demasiado pequeños.
* Elementos superpuestos.
* Layouts desktop comprimidos en móvil.

Verificación obligatoria

Después de cualquier cambio visual:

1. Ejecuta la aplicación.
2. Usa Playwright.
3. Captura desktop y móvil.
4. Revisa consola.
5. Prueba navegación.
6. Prueba botones.
7. Prueba el formulario sin enviar datos reales, salvo autorización.
8. Ejecuta las pruebas existentes.
9. Muestra un resumen de los archivos modificados.
10. Explica cualquier riesgo pendiente.

Restricciones técnicas

* No añadir dependencias sin necesidad.
* No reconstruir todo el proyecto para realizar un cambio localizado.
* No eliminar código porque parezca no utilizado sin verificar primero.
* No cambiar endpoints ni credenciales.
* No incluir secretos en el repositorio.
* No modificar configuración de despliegue sin autorización.
* No realizar cambios directos en producción.