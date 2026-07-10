# monselatte.com

Sitio público de **Monselatte Coffee Bar** (barra de café móvil para eventos en Puerto Rico). Su único objetivo es capturar leads de cotización.

## Arquitectura

```
Visitante → index.html (form #reserva)
              │  fetch POST (no-cors)
              ▼
    Google Apps Script Web App (doPost)
              │  valida + anti-spam + rate-limit
              ▼
    Google Sheet "Leads"  ──►  emails (admin + cliente)
              │
              │  dueño escribe `precio` y marca `aprobado` = si
              ▼
    trigger onSheetEdit/onLeadChange
              │  genera PDF (Drive /Cotizaciones) numerado ML-YYMM-###
              ▼
    email al cliente con PDF + link de Stripe (depósito $100)
              │  cliente paga → redirige a gracias.html
              ▼
    dueño marca `deposito` → PDF se reenvía como factura
```

- **Hosting**: GitHub Pages sirve la rama `main` desde la raíz. `CNAME` → monselatte.com. **Cada push a `main` publica al instante.**
- **Frontend**: HTML a mano + Tailwind CSS 3.4 (compilado localmente) + `main.js` vanilla (form, galería/lightbox, GA4).
- **Backend**: Google Apps Script ([apps-script.gs](apps-script.gs) es la copia de referencia; el código que corre vive en el editor de script.google.com).
- **Analytics**: GA4 `G-MGKWJKDS2L` en las 3 páginas, con eventos custom (`generate_lead`, `cta_click`, `contact`, `deposit_confirmed`).

## Desarrollo

```bash
npm install            # una vez
npm run build:css      # regenera assets/tw.css tras tocar HTML/tailwind.config.js/styles.css
npm run watch:css      # modo watch
```

- `styles.css` es el **input** de Tailwind (directivas + clases custom). El navegador solo carga `assets/tw.css` (compilado). **No enlazar `styles.css` en los HTML.**
- `tailwind.config.js` escanea `./*.html` y `./main.js` — si añades una página nueva `.html` en la raíz, queda cubierta automáticamente.
- **Siempre commitear `assets/tw.css` regenerado junto con los cambios de HTML** (no hay CI que lo compile).

## Apps Script — deploy (runbook)

El backend se despliega a mano (copy-paste). La URL del deployment vive en el `<meta name="sheet-webapp-url">` de `index.html` — `main.js` la lee de ahí.

1. Abrir el proyecto en [script.google.com](https://script.google.com) ("Monselatte Cotizaciones").
2. Pegar el contenido de `apps-script.gs` en `Code.gs`.
3. **Implementar → Administrar implementaciones → ✏️ editar la implementación existente → Versión: Nueva versión → Implementar.**
   Así la URL `/exec` **no cambia** y no hay que tocar el frontend.
   (Si se crea una "Nueva implementación" en su lugar, la URL cambia: actualizar los 2 meta tags.)
4. Probar: enviar una cotización desde el sitio y verificar fila nueva en el Sheet + emails.

### Script Properties (requeridas)

Los IDs sensibles NO están en el código (repo público). Se configuran una vez en
**Configuración del proyecto (⚙️) → Propiedades de la secuencia de comandos**:

| Propiedad | Contenido |
|---|---|
| `SPREADSHEET_ID` | ID del Google Sheet de leads |
| `LOGO_ID` | ID del archivo del logo en Drive |
| `QUOTES_FOLDER_ID` | ID de la carpeta /Cotizaciones en Drive |
| `TERMS_PDF_ID` | ID del PDF de términos y condiciones |
| `STRIPE_DEPOSIT_URL` | Payment Link del depósito (buy.stripe.com/...) |

Si falta alguna, el script falla al arrancar con un error que dice cuál.

### Triggers instalados (en el editor de Apps Script)

- `onSheetEdit` — instalable, "Al editar" el Spreadsheet (genera PDF/factura al aprobar o editar depósito).

## Estructura

| Ruta | Qué es |
|---|---|
| `index.html` | Landing one-page (hero, sobre, paquetes, galería, testimonios, form `#reserva`, FAQ) |
| `cotiza.html` | Redirect a `/#reserva` (página retirada; canonical a `/`) |
| `gracias.html` | Confirmación post-depósito (dispara GA4 `deposit_confirmed`) |
| `main.js` | Validación/envío del form, WhatsApp/email, galería, reveal, GA4 |
| `styles.css` | Input de Tailwind (no se sirve) |
| `assets/tw.css` | CSS compilado (el que se sirve) |
| `assets/img/optimized/` | Versiones .webp (los .jpeg quedan como fallback de `<picture>`) |
| `apps-script.gs` | Copia de referencia del backend |
| `sitemap.xml`, `robots.txt`, `CNAME` | SEO/dominio |

La copia vieja del sitio (`docs_old/`) vive en la rama `archive/docs_old`.

## Reglas de oro

- Push a `main` = producción inmediata. Verificar en preview local antes.
- No cambiar URLs (`/`, `/cotiza.html`, `/gracias.html`): Google las indexa y es el canal principal de clientes.
- No tocar el JSON-LD ni los meta tags SEO sin razón.
- El formulario es el negocio: cualquier cambio a `main.js`/form se prueba end-to-end (envío real de prueba) antes de push.
