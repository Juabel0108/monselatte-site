# Monselatte — Sistema visual aprobado

Documento del estado implementado en navbar, hero y «Planifica tu evento», inspeccionado el 23 de septiembre de 2026. No constituye un rediseño de las secciones restantes.

Fuentes: `index.html`, `styles.css`, `tailwind.config.js`, `main.js` y el CSS compilado `assets/tw.css`. Los valores siguientes corresponden al código actual. Los ejemplos en píxeles derivados de `rem` suponen la base habitual de 16 px del navegador; no se ha fijado otra base global.

## Propósito y personalidad

Monselatte presenta un servicio de barra de café móvil para bodas, eventos corporativos y celebraciones privadas en Puerto Rico. La web debe explicar el servicio, mostrar trabajo real, generar confianza y facilitar solicitudes de cotización y contacto.

La personalidad aprobada es editorial, elegante, cálida y refinada, vinculada a hospitality y specialty coffee. Debe transmitir atención personal y calidad de servicio, no una startup tecnológica ni una plantilla genérica.

## Principios

1. La fotografía real de Monselatte tiene prioridad sobre la decoración de UI.
2. No todo contenido debe vivir dentro de una card.
3. Whitespace, tipografía, fotografía y composición crean la jerarquía antes que sombras o containers.
4. Las cards se reservan principalmente para contenido realmente interactivo o cuando exista una razón semántica. La interacción tampoco obliga a usar una card.
5. Mobile debe diseñarse intencionalmente, no limitarse a apilar desktop.
6. No sacrificar SEO ni contenido útil para lograr minimalismo visual.
7. No modificar funcionalidad existente únicamente por razones estéticas.
8. Los futuros cambios deben reutilizar primero el sistema existente antes de introducir nuevos patrones.

Navbar y hero tienen dirección aprobada; no deben seguir rediseñándose salvo para evitar una regresión. La variante A es el hero definitivo. Las secciones inferiores y el formulario conservan por ahora su diseño anterior.

## Paleta

Tokens actuales de `theme.extend.colors.brand` en Tailwind:

| Token | Valor | Uso en el sistema aprobado |
| --- | --- | --- |
| `brand.green` | `#0B3D2E` | Marca, H1, números, navegación, botones y controles |
| `brand.cream` | `#F8F5F0` | Fondo continuo; texto sobre botones verdes |
| `brand.gold` | `#A57C2B` | Indicador de foco |
| `brand.text` | `#1A1A1A` | Texto base del body; equivalente conceptual a ink |
| `brand.espresso` | `#5A3A2E` | Descripción del hero, línea editorial y labels del planificador |

`ink` no es un token adicional definido en Tailwind. Las clases específicas aprobadas utilizan estos mismos colores mediante valores literales de CSS; no existe un sistema nuevo de variables CSS para estas tres áreas. No modificar los tokens globales para resolver ajustes locales.

## Tipografía

- **Instrument Serif**, peso 400: H1 del hero y números del planificador. Stack local: `'Instrument Serif', Georgia, serif`.
- **Inter**, pesos cargados 400, 500 y 600: body, navegación, botones y labels. Stack: `Inter, system-ui, sans-serif`.
- Google Fonts carga ambas familias con `display=swap` y conexiones anticipadas a sus dominios.
- **Playfair Display**, pesos 600 y 700, sigue cargada. El token global `fontFamily.serif` continúa siendo `['Playfair Display','serif']`. No se ha migrado el resto de los headings a Instrument Serif.

### Escala implementada

| Elemento | Base / móvil | Desde 640 px | Desde 1024 px | Peso; line-height; tracking |
| --- | --- | --- | --- | --- |
| Marca «Monselatte» | 16 px | Igual | Igual | 600; heredado; `-.025em` |
| Enlaces desktop | Ocultos | Ocultos | 13 px | 400 heredado; heredado; sin tracking propio |
| Enlaces del menú móvil | 15 px | Igual | Menú oculto | 400 heredado; heredado |
| CTA navbar | 13 px | Igual | Igual | 500; heredado |
| Símbolo hamburguesa | 22 px | Igual | Oculto | Sin peso/line-height propios |
| Eyebrow hero | 10 px | 11 px | 11 px | 500; `1.6`; `.14em` |
| H1 hero | `clamp(3.125rem, 6vw, 5.75rem)` | Misma fórmula | `clamp(3.25rem, 6vw, 5.75rem)` | 400; `.96`; `-.02em` |
| Descripción hero | 15 px | 16 px | 16 px | 400; `1.6` |
| CTA hero | 14 px | Igual | Igual | 500; heredado |
| Línea editorial inferior | 10 px | 10 px | 11 px | 400; `1.8`; `.11em` |
| Título y labels del planificador | 10 px | Igual | Igual | 500; `1.6`; `.13em` |
| Números del planificador | 52 px | Igual | Igual | 400; `1`; sin tracking propio |
| Botones +/− | 22 px | Igual | Igual | 400; `1` |
| CTA del planificador | 13 px | Igual | Igual | 500; heredado |
| Flechas de CTA hero/planificador | 18 px | Igual | Igual | Peso heredado; `1` |

El line-height base de Tailwind es `1.5`; se hereda donde no hay una regla local. No existe una escala global nueva que reemplace todos los tamaños del sitio.

Tamaños resultantes del H1, con base de 16 px:

| Viewport | H1 |
| --- | --- |
| 320, 375 y 390 px | 50 px |
| 768 px | 50 px |
| 1024 px | 61.44 px |
| 1440 px | 86.4 px |

El máximo de la fórmula es 92 px. El H1 no contiene saltos `<br>`. Desde 640 px tiene `max-width: 660px`; desde 1024 px se sustituye por `max-width: 5.9em`. En 1440 px esta restricción produce «Barra de café para / bodas y eventos / en Puerto Rico». Los saltos pueden variar con el viewport y la fuente disponible; no forzar dos líneas en móviles estrechos sacrificando legibilidad.

## Contenedores, gutters y spacing

Navbar, hero y planificador usan contenedores de **1440 px de ancho máximo**, centrados con `margin-inline: auto`. El máximo incluye padding bajo `box-sizing: border-box` de Tailwind. No equivale a un área útil de contenido de 1440 px.

| Área | Menos de 640 px | 640–1023 px | Desde 1024 px |
| --- | --- | --- | --- |
| Navbar: padding vertical / horizontal | 12 / 24 px | 12 / 40 px | 12 / 48 px |
| Navbar: altura mínima | 80 px | 80 px | 96 px |
| Hero: padding superior / lateral / inferior | 32 / 28 / 28 px | 48 / 40 / 32 px | 64 / 48 / 36 px |
| Hero: separación de grid | Filas: 36 px | 32 px | Columnas: 48 px; filas: 40 px |
| Planificador: padding vertical / horizontal | 28 / 28 px | 28 / 40 px | 26 / 48 px |
| Planificador: gap exterior | 24 px | 24 px | 40 px |
| Separación entre los dos selectores | 20 px | 32 px | `clamp(32px, 4vw, 64px)` |

Spacing interno del hero:

- Eyebrow → H1: `margin-top: 12px`; desde 1024 px, 16 px.
- H1 → descripción: 24 px; desde 1024 px, 28 px.
- Descripción → CTA: 24 px; desde 1024 px, 28 px.
- Ancho máximo de descripción: 350 px; desde 640 px, 440 px.
- La fotografía y la línea editorial se separan mediante el grid exterior, no con tarjetas.
- En desktop, la línea editorial ocupa ambas columnas y tiene padding superior de 24 px sobre su texto, después de su borde superior.

Navbar: gap de 24 px entre elementos principales; logo y nombre separados 12 px. Los enlaces desktop usan `gap: clamp(20px, 2.5vw, 36px)`.

## Breakpoints y composición responsive

Los cambios específicos del sistema aprobado utilizan `min-width: 640px` y `min-width: 1024px`. El cambio de navegación coincide con `lg` de Tailwind, 1024 px. La configuración no redefine los breakpoints de Tailwind. El CSS anterior también contiene un cambio a 768 px para `.cv`; no es un breakpoint adicional de composición de estas tres áreas.

### Hero

- Menos de 1024 px: una columna, orden texto y CTA → fotografía → línea editorial.
- Desde 1024 px: `grid-template-columns: minmax(0, 44fr) minmax(0, 56fr)`, con alineación centrada vertical. El 44/56 reparte el espacio disponible después de descontar el gap.
- Fotografía alineada al centro de su área con `align-self: center` y `min-width: 0`.
- Mantener presencia temprana de la fotografía en móvil. No agregar contenido antes de ella por decoración.
- H1 exacto: «Barra de café para bodas y eventos en Puerto Rico».
- Eyebrow exacto: «BARRA DE CAFÉ MÓVIL · PUERTO RICO».
- Descripción desktop: «Café de especialidad preparado al momento, servido desde nuestra barra móvil directamente en tu evento.»
- Descripción por debajo de 1024 px: «Café de especialidad preparado al momento, directamente en tu evento.»
- La diferencia de copy se implementa con un solo párrafo y un span: `.monselatte-hero__intro-detail` usa `display: none` y pasa a `inline` desde 1024 px. No son dos párrafos duplicados para lectores de pantalla.
- Línea editorial exacta: «BODAS · EVENTOS CORPORATIVOS · CELEBRACIONES PRIVADAS». Sigue siendo texto, no botones, pills ni cards.
- El hero conserva `scroll-mt-20`: margen de scroll de 5 rem, 80 px con la base habitual.

### Planifica tu evento

- Móvil: título, selector de invitados, selector de horas y CTA en composición vertical. El CTA ocupa el ancho disponible del grid.
- Desde 640 px: los selectores comparten una fila de dos columnas iguales con `repeat(2,minmax(0,1fr))`; título y CTA conservan sus filas. CTA alineado al inicio mediante `justify-self: start`.
- Desde 1024 px: título, bloque de controles y CTA en una fila; columnas `minmax(160px,1fr) minmax(360px,1.9fr) auto`.
- Es una transición compacta entre hero y contenido. No debe convertirse en otro hero ni prometer un precio calculado automáticamente.

Verificar siempre 320×568, 375×812, 390×844, 768×1024, 1024×768 y 1440×900. Evitar overflow horizontal, texto recortado, imágenes deformadas y solapamientos; probar también los puntos de cambio cuando se altere un breakpoint.

## Navbar y navegación

- Header `sticky`, `top: 0`, `z-index: 50`.
- Fondo sólido crema, sin desenfoque decorativo. `.site-header` y `.site-header.elevated` mantienen `box-shadow: none`.
- JavaScript sigue añadiendo `elevated` y `bg-brand-cream/95` cuando `scrollY > 10`; las reglas locales prevalecen y mantienen el tratamiento aprobado, sólido y sin sombra.
- Nombre visible: «Monselatte». Logo `assets/img/logo-oficial.png`, 48×48 px, `border-radius: 50%`. El logo circular es identidad, no un patrón para envolver iconos.
- Destinos desktop y móvil: Experiencia → `#sobre`; Eventos → `#paquetes`; Galería → `#galeria`; FAQ → `#faq`; Solicitar cotización → `#reserva`. Marca → `#inicio`.
- Desktop desde 1024 px; hamburguesa y menú móvil por debajo.
- Hamburguesa: 44×44 px. Conserva `menuBtn`, `mobileMenu`, `navLinks` y las clases funcionales existentes.
- `aria-expanded` y el nombre «Abrir menú»/«Cerrar menú» reflejan el estado. `aria-controls="mobileMenu"` identifica el panel.
- Cierra al seleccionar un enlace, con Escape y al cruzar el breakpoint de 1024 px. Escape devuelve el foco al botón.
- Panel móvil: `max-height: calc(100dvh - 80px)`, `overflow-y: auto`; padding interno `12px 24px 24px`. Enlaces con altura mínima de 48 px; CTA con margen superior de 12 px.
- Subrayado de hover/estado activo con offset de 6 px. El observador actual usa umbral `.25` y `rootMargin: '0px 0px -65% 0px'`. FAQ no forma parte de su lista observada: no asumir que todos los enlaces tienen resaltado activo automático.

## Botones y enlaces de acción

| Variante | Dimensiones y espaciado | Apariencia / hover |
| --- | --- | --- |
| CTA navbar | Mínimo 44 px de alto; padding `10px 18px` | Borde verde de 1 px, texto verde; hover verde con texto crema |
| CTA hero | Mínimo 48 px; padding `12px 20px`; gap 28 px | Fondo y borde verde de 1 px, texto crema; hover crema con texto verde |
| CTA planificador | Mínimo 48 px; padding `12px 18px`; gap 20 px | Transparente, borde verde de 1 px, texto verde; hover verde con texto crema |
| +/− planificador | 44×44 px | Transparente, borde verde al 25% de 1 px; hover verde con texto crema |

No hay sombras, pills, radios decorativos ni transformaciones de hover en estas variantes. Son controles rectangulares. Los CTA de hero y planificador usan una flecha tipográfica `→`, de 18 px, marcada `aria-hidden="true"`.

Mantener «Solicitar cotización» para navbar/hero y «Comenzar cotización» para el planificador. Todos conservan `href="#reserva"` y su tracking actual.

## Labels, líneas y divisores

Las etiquetas en mayúsculas son breves y funcionales. El hero utiliza `.14em` en su eyebrow; el planificador `.13em` en título y labels; la línea editorial del hero `.11em`. No extender estas mayúsculas a párrafos largos.

Todos los divisores siguientes tienen 1 px de grosor:

| Ubicación | Color |
| --- | --- |
| Borde inferior navbar y superior del menú móvil | `rgba(11,61,46,.14)` |
| Línea sobre tipos de evento, solo desktop | `rgba(11,61,46,.18)` |
| Bordes superior e inferior del planificador | `rgba(11,61,46,.16)` |
| Borde de +/− y subrayado de inputs numéricos | `rgba(11,61,46,.25)` |

Las líneas organizan y separan; no deben formar cajas alrededor de cada contenido. Los labels de invitados/horas tienen margen superior de 7 px y están centrados bajo el número. Ese centrado puntual no implica centrar todas las secciones.

## Fotografía

- Hero aprobado: **variante A**, `assets/img/optimized/galeria-4.webp`, con fallback `assets/img/galeria-4.jpeg`.
- Dimensiones declaradas y reales: 3424×2282 px. El WebP existente pesa aproximadamente 362 KiB.
- Render: `display: block`, `width: 100%`, `height: auto`, `aspect-ratio: 3 / 2`, `object-fit: cover`, `object-position: center`, `border-radius: 0`.
- Imagen integrada a la composición, sin sombra ni marco de card.
- Se utiliza `<picture>`, texto alternativo descriptivo, `loading="eager"`, `decoding="async"`, `fetchpriority="high"` y preload de la versión WebP.
- Preservar dimensiones declaradas, proporción y punto focal. No distorsionar ni ocultar detalles esenciales del servicio o de la marca.
- B y C fueron comparaciones temporales; no forman parte de un carrusel ni sustituyen la fotografía definitiva.
- Priorizar siempre fotografías reales de la barra, baristas, preparación y eventos de Monselatte.

## Controles de Planifica tu evento

Cada selector usa `grid-template-columns: 44px minmax(0,1fr) 44px`, gap de 12 px y alineación vertical centrada. Los campos y grids tienen `min-width: 0` para permitir ajuste responsive.

Los inputs tienen ancho 100%, altura 60 px, padding `0 2px`, fondo transparente, texto centrado, radio 0 y solo borde inferior. Sus números utilizan Instrument Serif 52 px/1, peso 400. Los spinners nativos se ocultan con `appearance: textfield` y las reglas WebKit; el input sigue siendo `type="number"` con `inputmode="numeric"` y conserva teclado y edición manual.

| Control | ID | Inicial | Mínimo | Máximo | Paso del input | Cambio mediante botones |
| --- | --- | --- | --- | --- | --- | --- |
| Invitados | `mini-invitados` | 50 | 1 | 500 | 1 | −10 / +10 |
| Horas | `mini-horas` | 2 | 2 | 8 | 1 | −1 / +1 |

Preservar `.mini-step`, `data-mini`, `data-delta`, `miniQuoteGo`, atributos y listeners. Los cambios manuales y botones conservan la lógica actual de límites; no sustituirla por reglas nuevas bajo una modificación visual.

El CTA copia los valores a `#leadForm [name="invitados"]` y `[name="horas_servicio"]`, dispara sus eventos `change` y navega a `#reserva`. Registra `cta_click` con `cta: 'mini_quote'`, `guests` y `service_hours`, además de la instrumentación general existente. Los inputs rápidos no tienen `name` y no añaden claves al payload. No son una calculadora de precios.

## Accesibilidad

- Conservar HTML semántico: navegación etiquetada, H1 real, enlaces para navegar y botones `type="button"` para +/−.
- La franja conserva `role="region"` y nombre accesible «Planifica tu evento».
- Labels asociados mediante `for`; nombres accesibles específicos para invitados, horas y cada botón +/−.
- Mantener acceso por Tab, activación por teclado, edición directa y flechas nativas de inputs numéricos.
- Targets de +/− y hamburguesa: 44×44 px; CTA hero y planificador: mínimo 48 px de alto. No reducirlos por estética.
- Focus navbar/hero: outline dorado de 2 px, offset de 5 px. Planificador: outline dorado de 2 px, offset de 4 px.
- No ocultar foco, depender exclusivamente del hover ni usar color como única indicación de estado.
- Mantener texto alternativo de imágenes y excluir las flechas decorativas de la lectura accesible.
- Preservar legibilidad, zoom y contraste al extender el sistema; estos valores documentan la implementación, no certifican una auditoría WCAG completa.

## Motion y reveal existentes

Navbar, hero y planificador aprobados no llevan `.reveal` ni animaciones nuevas. Sus estados de hover cambian color sin transiciones explícitas locales. No añadir movimiento para llenar espacios o llamar la atención sin una función.

El resto del sitio conserva:

- Scroll global suave: `html { scroll-behavior: smooth; }`.
- `.reveal`: opacidad 0, `translateY(12px)`, transiciones de opacidad y transform de `.6s ease`. `.reveal-visible` restaura opacidad 1 y elimina la transformación.
- IntersectionObserver de reveal con umbral `.12`; deja de observar cada elemento después de mostrarlo.
- Con `prefers-reduced-motion: reduce`, `.reveal` queda visible y sin transformación. El scroll suave global no tiene actualmente una excepción equivalente; no asumir que toda la página elimina movimiento con esa preferencia.
- Lightbox anterior: transición base `.2s`; animaciones de entrada/salida de fondo `.22s/.18s`, imagen `.26s/.18s` y caption `.24s/.16s`, con demora de `.06s` en la entrada del caption. Estas animaciones están dentro de `prefers-reduced-motion: no-preference`.
- FAQ anterior: transición de altura/opacidad `.3s` y giro del chevron `.25s`.

Estos efectos anteriores no son instrucciones para aplicarlos a nuevas secciones. Cualquier nueva animación deberá tener una razón funcional y respetar reducción de movimiento.

## Patrones que debemos evitar

- **Excessive cards**: envolver cada bloque de contenido en una tarjeta.
- **Bento grids sin justificación** semántica o funcional.
- **Glassmorphism** y superficies translúcidas decorativas.
- **Gradient blobs** y gradientes decorativos.
- **Decorative icon circles**: círculos e iconos que no aportan función.
- **Excessive pills**.
- **Excessive border-radius**.
- **Shadows decorativas**, pronunciadas o usadas como jerarquía principal.
- **Generic feature grids** que sustituyan composición y contenido específicos.
- **Centrar todas las secciones**.
- **Stock photography**.
- **AI imagery**.
- **Animaciones innecesarias**, glow y transformaciones llamativas.
- Apariencia SaaS/startup genérica, nuevos colores sin motivo o elementos añadidos solo para llenar espacio.

El repositorio aún contiene sombras, radios, pills, gradientes y cards en secciones anteriores al sistema aprobado. Por ejemplo, Tailwind conserva `shadow.soft: 0 10px 30px rgba(0,0,0,.08)` y `borderRadius.xl2: 1rem`. Su existencia no los convierte en patrones aprobados para extender el rediseño ni autoriza eliminarlos fuera del alcance solicitado.

## Reutilización y protección funcional

Reutilizar las familias actuales `.site-*`, `.monselatte-hero__*` y `.event-planner__*` como referencia de valores y comportamiento. No cambiar selectores compartidos, fuentes globales o colores globales para resolver una sola sección. Tampoco crear un patrón nuevo cuando el existente cubra la necesidad.

`styles.css` es la fuente; el navegador carga `assets/tw.css`, generado con `npm run build:css`. No editar manualmente el compilado ni enlazar el CSS fuente en la página.

Conservar contenido SEO, H1, IDs, destinos, tracking y funcionalidad existente. Antes de intervenir archivos relacionados con cotizaciones, seguir `AGENTS.md` y `docs/INTEGRATIONS.md`. No cambiar formulario, endpoint, payload, Apps Script, validaciones o mensajes por razones estéticas. No efectuar envíos reales durante pruebas sin autorización.

Los cambios visuales futuros deben comprobar responsive, teclado, navegación, consola y regresiones funcionales con Playwright y las pruebas existentes, limitándose a la sección autorizada. Este documento no autoriza cambios adicionales ni publicación.
