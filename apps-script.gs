/************************************************************
 * Monselatte — Web App (limpio, sin Google Docs/PDF)
 * - Guarda leads en Google Sheets
 * - Validaciones básicas y anti-spam
 * - Rate limit (5 envíos / 10 min por email/teléfono)
 * - Envía correos (cliente + admin) con diseño simple
 ************************************************************/

// ====== CONFIG ======
const SPREADSHEET_ID = '1Y0Dcn9x26myWrG7b8E49Q1AyX7Nfg2fBKG_lElCX5uI';
const SHEET_NAME     = 'Leads';    // la hoja donde guardamos
const SPAM_SHEET     = 'Spam';     // hoja para registros descartados
const LOGO_ID = '1ujLjOOxvs-QWuU-OwtdXXbaoGmEa1BEQ';
// === PDF / Carpeta y Términos ===
const QUOTES_FOLDER_ID = '1s56ccSMJ2UCXxTWxCj2PuLV29rZPhcsQ';
const TERMS_PDF_ID     = '16gKZrQ2viHr3kP2QZPlsHsPomTavIAgY';

const ADMIN_EMAIL    = 'monselattepr@gmail.com';
const COMPANY = {
  name: 'Monselatte Coffee Bar',
  phone: '(787) 610-8953',
  email: 'monselattepr@gmail.com',
  site: 'monselatte.com'
};

const ATTACH_QUOTE_PDF = true; // si ves doble adjunto, ponlo en false para enviar solo el link
const INCLUDE_DRIVE_LINK = false; // enviar SOLO el PDF adjunto (sin link de Drive)

const FIXED_DEPOSIT_AMOUNT = 100;
const STRIPE_DEPOSIT_URL   = 'https://buy.stripe.com/6oU9AV4Zw5SLcMmbjp7ok00';
const THANK_YOU_URL        = 'https://monselatte.com/gracias.html';

/** Devuelve el logo como blob para inlineImages (cid:logo). Asegúrate de poner tu LOGO_ID. */
function getLogoBlob() {
  try {
    return DriveApp.getFileById(LOGO_ID).getBlob().setName('logo.png');
  } catch (e) {
    // Fallback: crea un blob vacío si falla (evita romper el envío)
    return Utilities.newBlob('', 'image/png', 'logo.png');
  }
}

// Orden y nombre de columnas a crear si no existen (puedes ajustar)
const LEADS_HEADERS = [
  'timestamp','nombre','email','telefono',
  'fecha','hora_inicio','horas_servicio','hora_fin',
  'localidad','direccion',
  'tipo','invitados','paquete','mensaje',

  // ✅ en este orden (como tu Sheet)
  'precio','aprobado','deposito',

  // tracking
  'utm_source','utm_medium','utm_campaign','utm_term','utm_content',
  'referrer','page_path',

  // pdf workflow
  'nro','pdfUrl','sent_at'
];

// ====== WEB APP ======
function doPost(e) {
  const res = (obj) => ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  // 0) Body presente
  if (!e || !e.postData || !e.postData.contents) {
    return res({ ok:false, error:'no-body' });
  }

  // 1) Parse body
  let data;
  try {
    data = JSON.parse(e.postData.contents || '{}');
  } catch (parseErr) {
    return res({ ok: false, error: 'invalid-json' });
  }

  // 2) HONEYPOT: si viene website -> SPAM (no guardamos en Leads)
  if ((data.website || '').trim() !== '') {
    logSpam('honeypot', data);
    return res({ ok:false, reason:'honeypot' });
  }

  // 3) Validaciones
  const errors = validate(data);
  if (errors.length) {
    logSpam('validation: ' + errors.join(' | '), data);
    return res({ ok:false, reason:'validation', errors });
  }

  // 4) Rate limit
  if (isRateLimited(data)) {
    logSpam('rate-limit', data);
    return res({ ok:false, reason:'rate-limit' });
  }

  // 5) Guardar en Leads
  const ss = getOrCreateSheet(SPREADSHEET_ID, SHEET_NAME, LEADS_HEADERS);
  const clean = sanitizeInput(data);

ss.appendRow([
  new Date(),
  clean.nombre,
  clean.email,
  clean.telefono,
  clean.fecha,

  clean.hora_inicio,
  clean.horas_servicio,
  clean.hora_fin,

  clean.localidad,
  clean.direccion,
  clean.tipo,
  clean.invitados,
  clean.paquete,
  clean.mensaje,

  '', '', '', // ✅ precio, aprobado, deposito

  (data.utm_source  || ''), (data.utm_medium  || ''), (data.utm_campaign || ''),
  (data.utm_term    || ''), (data.utm_content || ''),
  (data.referrer    || ''), (data.page_path   || ''),

  '', '', ''
]);

  // 6) Emails
  try {
    sendAdminEmail(clean);
  } catch (err1) { console.warn('Admin email fail:', err1); }

  try {
    if ((clean.email || '').trim() !== '') {
      sendClientEmail(clean);
    }
  } catch (err2) { console.warn('Client email fail:', err2); }

  return res({ ok:true });
}

// ====== VALIDACIONES ======
function validate(d) {
  const errors = [];
  const reEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const invitados = parseInt(d.invitados || '0', 10);

  // nombre
  if (!d.nombre || !/^[A-Za-zÁÉÍÓÚÑáéíóúñ'’\-\s]{2,60}$/.test(d.nombre.trim())) errors.push('nombre');

  // email
  if (!reEmail.test(d.email || '')) errors.push('email');

  // telefono (>=7 y <=15 dígitos)
  const tel = digits(d.telefono || '');
  if (tel.length < 7 || tel.length > 15) errors.push('telefono');

  // fecha (opcionalmente validamos formato ISO yyyy-mm-dd)
  if (!d.fecha) errors.push('fecha');

  // hora_inicio requerida
  if (!d.hora_inicio) errors.push('hora_inicio');

  // hora_fin requerida (para opción C)
  if (!d.hora_fin) errors.push('hora_fin');

  // horas_servicio: aceptamos que venga del frontend, pero si no viene lo calculamos con inicio+fin.
  const hs = parseInt(d.horas_servicio || '0', 10);
  const hsCalc = calcHorasServicio(d.hora_inicio, d.hora_fin);
  const hsFinal = (hs > 0 ? hs : hsCalc);
  if (!(hsFinal >= 1 && hsFinal <= 12)) errors.push('horas_servicio');

  // invitados razonables
  if (!(invitados >= 1 && invitados <= 500)) errors.push('invitados');

  // localidad PR
  if (!d.localidad) {
    errors.push('localidad_requerida');
  } else if (!PUERTO_RICO_MUNICIPIOS.has(String(d.localidad))) {
    errors.push('localidad_invalida');
  }

  return errors;
}

// ====== RATE LIMIT ======
function isRateLimited(data) {
  const cache = CacheService.getScriptCache();
  const key = ('lead:' + (data.email || '').toLowerCase() + ':' + digits(data.telefono || '')).slice(0,240);
  const prev = cache.get(key);
  if (prev) {
    const count = parseInt(prev, 10) || 0;
    if (count >= 5) return true;  // 5 envíos en 10 min
    cache.put(key, String(count + 1), 600);
    return false;
  }
  cache.put(key, '1', 600); // 10 minutos
  return false;
}

// ====== HELPERS ======
function getOrCreateSheet(spreadsheetId, sheetName, headers) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  // Headers
  const firstRow = sh.getRange(1,1,1, sh.getLastColumn() || headers.length).getValues()[0];
  if (!firstRow || firstRow.join('') === '') {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}


function sanitizeInput(d) {
  const horaInicio12 = toHora12(d.hora_inicio || '');
  const horaFin12    = toHora12(d.hora_fin || '') || calcHoraFin(d.hora_inicio, d.horas_servicio);

  // Opción C: si no llega horas_servicio, la calculamos desde inicio+fin
  const hsCalc = calcHorasServicio(d.hora_inicio, d.hora_fin);
  const hsFinal = (String(d.horas_servicio || '').trim() !== '' ? parseInt(d.horas_servicio, 10) : hsCalc);

  return {
    nombre:   safe(d.nombre),
    email:    safe(d.email),
    telefono: digits(d.telefono || ''),
    fecha:    safe(d.fecha),

    hora_inicio:    horaInicio12,
    horas_servicio: String(isFinite(hsFinal) && hsFinal > 0 ? hsFinal : ''),
    hora_fin:       horaFin12,

    localidad: safe(d.localidad),
    direccion: safe(d.direccion),
    tipo:      safe(d.tipo),
    invitados: String(parseInt(d.invitados || '0', 10)),
    paquete:   safe(d.paquete),
    mensaje:   safe(d.mensaje)
  };
}

function digits(s) { return String(s || '').replace(/[^\d]/g,''); }
function safe(v)   { return v == null ? '' : String(v).trim(); }

// Normaliza "Sí", "si", checkbox TRUE, etc.
function normalizeYes_(v) {
  const s = String(v ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, ''); // quita acentos

  return (
    s === 'si' ||
    s === 'yes' ||
    s === 'true' ||
    s === '1'
  );
}

// Hora a 12h (intenta reconocer 24h o 12h confusa)
function toHora12(hstr) {
  const s = String(hstr || '').trim();
  if (!s) return '';
  // ya viene con am/pm?
  const mAm = s.toLowerCase().includes('am');
  const mPm = s.toLowerCase().includes('pm');
  if (mAm || mPm) {
    // normalizar espacios, Mayúsculas
    return s.replace(/\s+/g,'').toUpperCase().replace(/AM|PM/, mAm ? ' AM' : ' PM');
  }
  // viene en 24h? (HH:mm)
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return s; // lo dejamos tal cual
  let hh = parseInt(m[1],10), mm = m[2];
  let suf = ' AM';
  if (hh === 0)      { hh = 12; suf = ' AM'; }
  else if (hh === 12){ suf = ' PM'; }
  else if (hh > 12)  { hh = hh - 12; suf = ' PM'; }
  return `${hh}:${mm}${suf}`;
}

// Calcula hora_fin si el frontend no la envía: hora_inicio + horas_servicio
function calcHoraFin(horaInicio, horasServicio) {
  const hs = parseInt(horasServicio || '0', 10);
  if (!horaInicio || !(hs >= 1 && hs <= 12)) return '';

  const s = String(horaInicio).trim();
  if (!s) return '';

  // convertir a minutos desde 00:00; aceptamos "HH:mm" o "H:mm AM/PM"
  let hh, mm;
  const hasAmPm = s.toLowerCase().includes('am') || s.toLowerCase().includes('pm');

  if (hasAmPm) {
    const isPM = s.toLowerCase().includes('pm');
    const isAM = s.toLowerCase().includes('am');
    const m = s.match(/(\d{1,2}):(\d{2})/);
    if (!m) return '';
    hh = parseInt(m[1], 10);
    mm = parseInt(m[2], 10);
    if (isPM && hh !== 12) hh += 12;
    if (isAM && hh === 12) hh = 0;
  } else {
    const m = s.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return '';
    hh = parseInt(m[1], 10);
    mm = parseInt(m[2], 10);
  }

  let total = hh * 60 + mm + hs * 60;
  total = total % (24 * 60);

  const endH = Math.floor(total / 60);
  const endM = String(total % 60).padStart(2, '0');
  return toHora12(`${endH}:${endM}`);
}

// Calcula horas_servicio (entero) usando hora_inicio y hora_fin.
// Devuelve 0 si no puede calcular.
function calcHorasServicio(horaInicio, horaFin) {
  const startMin = parseHoraToMinutes_(horaInicio);
  const endMin   = parseHoraToMinutes_(horaFin);
  if (startMin < 0 || endMin < 0) return 0;

  // Si fin es "menor" que inicio, asumimos que cruza medianoche.
  let diff = endMin - startMin;
  if (diff <= 0) diff += 24 * 60;

  // Redondeo: a horas enteras. Si necesitas medias horas, lo ajustamos luego.
  const hours = Math.round(diff / 60);
  return (hours >= 1 && hours <= 24) ? hours : 0;
}

// Convierte "HH:mm" o "H:mm AM/PM" a minutos desde 00:00.
// Devuelve -1 si falla.
function parseHoraToMinutes_(val) {
  const s = String(val || '').trim();
  if (!s) return -1;

  const lower = s.toLowerCase();
  const hasAm = lower.includes('am');
  const hasPm = lower.includes('pm');

  // Captura HH:mm
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (!m) return -1;

  let hh = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  if (!(mm >= 0 && mm <= 59) || !(hh >= 0 && hh <= 23)) return -1;

  // Si viene en AM/PM, normalizamos a 24h
  if (hasAm || hasPm) {
    if (hh === 12) hh = 0;
    if (hasPm) hh += 12;
    if (hh < 0 || hh > 23) return -1;
  }

  return hh * 60 + mm;
}

// ====== EMAILS ======
function sendAdminEmail(lead) {
  const subject = 'Nueva cotización — monselatte.com';
  const {html, text} = composeEmail(lead, {forAdmin:true});
  GmailApp.sendEmail(
    ADMIN_EMAIL,
    subject,
    text,
    { name: COMPANY.name, htmlBody: html, replyTo: lead.email || COMPANY.email, inlineImages: { logo: getLogoBlob() } }
  );
}

function sendClientEmail(lead) {
  const subject = 'Monselatte recibió tu solicitud';
  const {html, text} = composeEmail(lead, {forAdmin:false});
  GmailApp.sendEmail(
    lead.email,
    subject,
    text,
    { name: COMPANY.name, htmlBody: html, replyTo: COMPANY.email, inlineImages: { logo: getLogoBlob() } }
  );
}

function composeEmail(lead, {forAdmin}) {
  // Campos listados
const fields = [
  ['Fecha',         lead.fecha || '—'],
  ['Hora inicio',   lead.hora_inicio || lead.hora || '—'],
  ['Horas',         lead.horas_servicio || '—'],
  ['Hora fin',      lead.hora_fin || '—'],
  ['Municipio',     lead.localidad || '—'],
  ['Dirección',     lead.direccion || '—'],
  ['Tipo de evento',lead.tipo || '—'],
  ['Invitados',     lead.invitados || '—']
];

  // Filas adicionales para admin (contacto)
  const adminExtra = forAdmin ? [
    ['Nombre',     lead.nombre || '—'],
    ['Email',      lead.email || '—'],
    ['Teléfono',   lead.telefono || '—']
  ] : [];

  // Texto simple (fallback)
  let text = '';
  if (forAdmin) {
    text =
`Se recibió una nueva solicitud:

Nombre: ${lead.nombre}
Email: ${lead.email}
Teléfono: ${lead.telefono}
Fecha: ${lead.fecha}
Hora inicio: ${lead.hora_inicio || lead.hora}
Horas: ${lead.horas_servicio || '—'}
Hora fin: ${lead.hora_fin || '—'}
Localidad: ${lead.localidad}
Dirección: ${lead.direccion}
Tipo de evento: ${lead.tipo}
Invitados: ${lead.invitados}
Mensaje: ${lead.mensaje || '—'}
`;
  } else {
    text =
`¡Hola ${lead.nombre}!

Gracias por contactarnos en ${COMPANY.name}.
Te responderemos en 24–48 h.

Resumen:
Fecha: ${lead.fecha}
Hora inicio: ${lead.hora_inicio || lead.hora}
Horas: ${lead.horas_servicio || '—'}
Hora fin: ${lead.hora_fin || '—'}
Localidad: ${lead.localidad}
Dirección: ${lead.direccion}
Tipo de evento: ${lead.tipo}
Invitados: ${lead.invitados}

— ${COMPANY.name}`;
  }

  // HTML con diseño simple
  const rows = [...adminExtra, ...fields].map(([k,v]) => `
    <tr>
      <td style="padding:10px 12px;border:1px solid #eee;background:#fafafa;width:40%;font-weight:600">${k}</td>
      <td style="padding:10px 12px;border:1px solid #eee">${escapeHtml(v)}</td>
    </tr>
  `).join('');

  const topMsg = forAdmin
    ? `Se recibió una nueva solicitud:`
    : `Aquí tienes los detalles de tu solicitud. Te responderemos en 24–48 h.`;

  const introHi = forAdmin
    ? ''
    : `¡Hola <strong>${escapeHtml(lead.nombre || '')}</strong>!`;

  const extraMsg = forAdmin
    ? `<p style="margin:0;color:#6b7280;font-size:14px">Mensaje: ${escapeHtml(lead.mensaje || '—')}</p>`
    : '';

  const html = `
  <div style="background:#f6f7f9;padding:24px 0;font-family:Inter,system-ui,Segoe UI,Roboto,Arial,sans-serif">
    <div style="max-width:680px;margin:0 auto;background:#fff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">
      <div style="display:flex;gap:16px;align-items:center;padding:16px 20px;border-bottom:1px solid #eee">
        <img src="cid:logo" alt="${COMPANY.name}" width="44" height="44" style="border-radius:10px;display:block"/>
        <div style="line-height:1.3">
          <div style="font-weight:700">${COMPANY.name}</div>
          <div style="font-size:12px;color:#6b7280">
            Barra de café para eventos en Puerto Rico<br/>
            Tel: ${COMPANY.phone} &nbsp;·&nbsp;
            <a href="mailto:${COMPANY.email}" style="color:#6b7280">${COMPANY.email}</a>
          </div>
        </div>
      </div>

      <div style="padding:16px 20px">
        ${introHi}
        <p style="margin:8px 0 16px 0;color:#111827">${topMsg}</p>

        <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #eee;border-radius:8px;overflow:hidden">
          ${rows}
        </table>

        ${extraMsg}

        ${forAdmin ? '' : `
        <div style="margin-top:18px;color:#6b7280;font-size:13px">
          Si no recibes respuesta en 24–48 h, escríbenos por WhatsApp.
        </div>`}
      </div>

      <div style="border-top:1px solid #eee;padding:14px 20px;background:#fafafa;color:#6b7280;font-size:12px">
        © ${new Date().getFullYear()} ${COMPANY.name} · <a href="https://${COMPANY.site}" style="color:#6b7280">${COMPANY.site}</a>
      </div>
    </div>
  </div>`.trim();

  return { html, text };
}

function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

// ====== SPAM LOG ======
function logSpam(reason, data) {
  try {
    const sh = getOrCreateSheet(SPREADSHEET_ID, SPAM_SHEET, ['timestamp','reason','payload']);
    sh.appendRow([
      new Date(),
      reason,
      typeof data === 'string' ? data : JSON.stringify(data)
    ]);
  } catch (e) {
    console.warn('No se pudo escribir en Spam:', e);
  }
}

// ====== Municipios de Puerto Rico ======
const PUERTO_RICO_MUNICIPIOS = new Set([
  'Adjuntas','Aguada','Aguadilla','Aguas Buenas','Aibonito','Arecibo','Arroyo',
  'Barceloneta','Barranquitas','Bayamón','Cabo Rojo','Caguas','Camuy','Canóvanas',
  'Carolina','Cataño','Cayey','Ceiba','Ciales','Cidra','Coamo','Comerío','Corozal',
  'Culebra','Dorado','Fajardo','Florida','Guánica','Guayama','Guayanilla','Guaynabo',
  'Gurabo','Hatillo','Hormigueros','Humacao','Isabela','Jayuya','Juana Díaz',
  'Juncos','Lajas','Lares','Las Marías','Las Piedras','Loíza','Luquillo','Manatí',
  'Maricao','Maunabo','Mayagüez','Moca','Morovis','Naguabo','Naranjito','Orocovis',
  'Patillas','Peñuelas','Ponce','Quebradillas','Rincón','Río Grande','Sabana Grande',
  'Salinas','San Germán','San Juan','San Lorenzo','San Sebastián','Santa Isabel',
  'Toa Alta','Toa Baja','Trujillo Alto','Utuado','Vega Alta','Vega Baja','Vieques',
  'Villalba','Yabucoa','Yauco'
]);

/************************************************************
 * Cómo desplegar:
 * 1) Publicar > Implementar como aplicación web
 * 2) Nueva implementación > Ejecutar como tú > Acceso: Cualquiera con el link
 * 3) Reemplaza en tu frontend el endpoint con la URL de despliegue
 *
 * Body esperado (JSON):
 * {
"nombre","email","telefono","fecha","hora_inicio","horas_servicio","hora_fin",
 *   "localidad","direccion","tipo","invitados","paquete","mensaje",
 *   "website" // honeypot invisible, debe llegar vacío
 *   // Opcionales de tracking (si existen en la URL o el frontend):
 *   "utm_source","utm_medium","utm_campaign","utm_term","utm_content",
 *   "referrer","page_path"
 * }
 ************************************************************/

/** Genera número corto de cotización por mes: ML-YYMM-### (ej. ML-2509-007) */
function buildQuoteNumber() {
  const yymm = getYYMM_();
  const seq = getNextMonthlySeq_(yymm);
  return `ML-${yymm}-${seq}`;
}

// Devuelve YYMM actual
function getYYMM_() {
  const d = new Date();
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return yy + mm;
}

// Lleva el consecutivo en PropertiesService por cada YYMM
function getNextMonthlySeq_(yymm) {
  const props = PropertiesService.getScriptProperties();
  const key = `SEQ_${yymm}`;
  let n = Number(props.getProperty(key) || '0');
  n += 1;
  props.setProperty(key, String(n));
  return String(n).padStart(3, '0');
}

/** Mapea la fila a variables para la plantilla */
function mapRowForTemplate(r) {
  const moneyFmt = new Intl.NumberFormat('es-PR', { style:'currency', currency:'USD', minimumFractionDigits:2 });
  const money = (n)=> (isFinite(n) ? moneyFmt.format(Number(n)) : '—');
  const precio   = Number(r.precio || 0);
  const depositoPagado   = Number(r.deposito || 0);
  const depositoMostrado = depositoPagado > 0 ? depositoPagado : FIXED_DEPOSIT_AMOUNT;
  const balance          = Math.max(precio - depositoMostrado, 0);
  const depositoSugerido = FIXED_DEPOSIT_AMOUNT;
  const isInvoice        = depositoPagado > 0;
  const docTitle = isInvoice
    ? 'FACTURA DE BARRA DE CAFÉ'
    : 'COTIZACIÓN PARA SERVICIOS DE BARRA DE CAFÉ';
  const docLabel = isInvoice ? 'FACTURA' : 'COTIZACIÓN';
  const docSubtitle = 'Servicios de barra de café';

  // --- FECHA BONITA (ES) ---
  const tz = 'America/Puerto_Rico';
  const MESES_ES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  let fechaBonita = '';
  try {
    // la celda puede venir como Date, número (serial), o string
    let d = r.fecha instanceof Date ? r.fecha : new Date(r.fecha);
    if (r.fecha && typeof r.fecha === 'number') d = new Date(Math.round((r.fecha - 25569) * 86400 * 1000)); // por si fuera serial XLS
    if (!isNaN(d)) {
      const dd = Utilities.formatDate(d, tz, 'd');
      const mm = MESES_ES[d.getMonth()];
      const yyyy = Utilities.formatDate(d, tz, 'yyyy');
      fechaBonita = `${dd} de ${mm} de ${yyyy}`;
    }
  } catch (e) {
    fechaBonita = String(r.fecha || '');
  }

  // --- FECHA CORTA (ES) ---
  // Ej: "31 ene 2026" (evita que la fecha se parta en dos líneas)
  const MESES_CORTO = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  let fechaCorta = '';
  try {
    let d2 = r.fecha instanceof Date ? r.fecha : new Date(r.fecha);
    if (r.fecha && typeof r.fecha === 'number') d2 = new Date(Math.round((r.fecha - 25569) * 86400 * 1000));
    if (!isNaN(d2)) {
      const dd2 = Utilities.formatDate(d2, tz, 'd');
      const mm2 = MESES_CORTO[d2.getMonth()];
      const yyyy2 = Utilities.formatDate(d2, tz, 'yyyy');
      fechaCorta = `${dd2} ${mm2} ${yyyy2}`;
    }
  } catch (e) {
    fechaCorta = '';
  }

  // --- HORA BONITA (12h) ---
  function hora12FromAny(val) {
    // Si viene como Date (hora), formatea; si viene como string/numero, normaliza a 12h.
    try {
      if (val instanceof Date) {
        return Utilities.formatDate(val, tz, 'h:mm a');
      }
      if (typeof val === 'number') {
        // posible serial de Excel (parte fraccionaria del día)
        const ms = Math.round(val * 24 * 60 * 60 * 1000);
        const d = new Date(ms);
        const h = Utilities.formatDate(d, tz, 'h:mm a');
        return h;
      }
      // texto: reusa tu normalizador
      return toHora12(String(val || ''));
    } catch (e) {
      return String(val || '');
    }
  }
const horaInicioBonita = hora12FromAny(r.hora_inicio || r.hora);
const horaFinBonita    = hora12FromAny(r.hora_fin);

  // --- INCLUDES POR PAQUETE ---
const includes = getPackageIncludes_();

  // --- LOGO COMO DATA URL PNG (forzado a PNG para que renderice) ---
  let logo = '';
  try {
    const blobPng = getLogoBlob().getAs('image/png');
    const b64 = Utilities.base64Encode(blobPng.getBytes());
    logo = 'data:image/png;base64,' + b64;
  } catch (e) { logo = ''; }

  return {
    nro:        r.nro || '',
    fecha:      r.fecha || '',
    hora:        horaInicioBonita || '',
    hora_inicio: horaInicioBonita || '',
    hora_fin:    horaFinBonita || '',
    horas_servicio: r.horas_servicio || '',
    nombre:     r.nombre || '',
    email:      r.email || '',
    telefono:   r.telefono || '',
    localidad:  r.localidad || '',
    direccion:  r.direccion || '',
    tipo:       r.tipo || '',
    invitados:  r.invitados || '',
    paquete:    r.paquete || '',
    precio:     money(precio),
    deposito:   money(depositoMostrado),
    balance:    money(balance),
    balance50:  money(depositoSugerido),
    company:    { name: COMPANY.name, phone: COMPANY.phone, email: COMPANY.email },
    fechaBonita: fechaBonita || String(r.fecha || ''),
    fechaCorta:  fechaCorta || '',
    includes:   includes,
    isInvoice:  isInvoice,
    docTitle:   docTitle,
    docLabel:   docLabel,
    docSubtitle: docSubtitle,
    logo:       logo
  };
}

function getPackageIncludes_() {
  return [
    'Baristas preparando café en el momento',
    'Máquina de espresso premium',
    'Menú: latte, cappuccino, espresso, americano y cortado',
    'Leche regular + alternativas (según disponibilidad)',
    'Sabores: vainilla, caramelo y temporada',
    'Vasos, tapas, servilletas, azúcar y montaje completo'
  ];
}

/** HTML base del PDF (ajústalo a tu branding) */
function quoteTemplateHtml() {
  return `
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title><?= data.docTitle ?></title>
<style>
  *{ box-sizing:border-box; font-family: Inter, Arial, sans-serif; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body{ margin:0; padding:24px; color:#111; }
  .card{ border:1px solid #E5E7EB; border-radius:16px; overflow:hidden }
  .header{ display:flex; gap:16px; align-items:center; padding:16px 20px; background:#F3F4F6; }
  .brand{ font-weight:800; font-size:16px }
  .muted{ color:#6B7280; font-size:12.5px }
  .docLabel{ font-size:11px; font-weight:800; color:#6B7280; letter-spacing:1px; text-transform:uppercase }
  .docTitle{ margin:0; font-size:22px; line-height:1.2; letter-spacing:-0.2px; font-weight:800; color:#111827 }
  .badge{ display:inline-block; background:#214d45; color:#fff; padding:4px 10px; border-radius:999px; font-weight:800; font-size:12px; letter-spacing:0.4px; vertical-align:middle }
  .metaBlock{ font-weight:700; text-align:right; line-height:1.35; padding-top:2px }
  .metaLine{ white-space:nowrap }
  table{ width:100%; border-collapse:collapse; margin-top:12px }
  th, td{ border:1px solid #E5E7EB; padding:10px 12px; vertical-align:top }
  th{ text-align:left }
  .green{ background:#214d45; color:#fff }
  .totals td{ font-weight:700 }
</style>
</head>
<body>
  <div class="card">
    <div class="header">
      <img src="<?!= data.logo ?>" alt="logo" width="48" height="48" style="border-radius:10px;object-fit:contain"/>
      <div>
        <div class="brand">Monselatte Coffee Bar</div>
        <div class="muted">Barra de café para eventos en Puerto Rico · Tel: <?= data.company.phone ?> · <?= data.company.email ?></div>
      </div>
    </div>

    <div style="padding:16px 20px">
  <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:16px">
    <div style="min-width:340px">
      <div class="docLabel"><?= data.docLabel ?></div>
      <div class="docTitle">Servicios de barra de café</div>
      <div style="margin-top:6px">
        <span class="badge"><?= data.nro ?></span>
      </div>
    </div>

    <div class="muted metaBlock">
      <div class="metaLine">
        <span style="color:#111827">Fecha del evento:</span>
        <?= data.fechaCorta || data.fechaBonita || data.fecha ?>
      </div>
      <div class="metaLine" style="margin-top:2px">
        <span style="color:#111827">Horario:</span>
        <?= data.hora_inicio ?> – <?= data.hora_fin ?>
      </div>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:12px">
        <div>
          <table>
            <tr><th style="width:40%">Cliente</th><td><?= data.nombre ?></td></tr>
            <tr><th>Email</th><td><?= data.email ?></td></tr>
            <tr><th>Teléfono</th><td><?= data.telefono ?></td></tr>
          </table>
        </div>
        <div>
          <table>
            <tr><th style="width:40%">Municipio</th><td><?= data.localidad ?></td></tr>
            <tr><th>Dirección</th><td><?= data.direccion ?></td></tr>
            <tr><th>Tipo de evento</th><td><?= data.tipo ?></td></tr>
            <tr><th>Invitados</th><td><?= data.invitados ?></td></tr>
          </table>
        </div>
      </div>

      <table style="margin-top:16px">
        <tr>
          <th class="green" style="width:28%">Barra de café</th>
          <th class="green">Incluye:</th>
          <th class="green" style="width:18%">Precio</th>
        </tr>
        <tr>
          <td><?= data.paquete ?></td>
          <td>
            <ul style="margin:0;padding-left:16px">
             <? for (var i=0; i < data.includes.length; i++) { ?>
              <li style="margin:0 0 4px 0;line-height:1.25"><?!= data.includes[i] ?></li>
             <? } ?>
          </ul>
          </td>
          <td style="text-align:right;font-weight:700"><?= data.precio ?></td>
        </tr>
      </table>

      <table style="margin-top:12px">
        <tr><th>Depósito para reservar</th><td><?= data.deposito ?></td></tr>
        <tr class="totals"><th>Balance restante</th><td><?= data.balance ?></td></tr>
      </table>

      <? if (data.isInvoice) { ?>
      <div class="muted" style="margin-top:10px">
        <strong>Notas:</strong> Depósito recibido: <?= data.deposito ?>.<br/>
        El balance pendiente (<?= data.balance ?>) se paga el día del evento antes de comenzar el servicio.
      </div>
      <? } else { ?>
      <div class="muted" style="margin-top:10px">
        <strong>Notas:</strong> Se requiere un depósito fijo de <?= data.deposito ?> para separar la fecha.<br/>
        El balance restante (<?= data.balance ?>) debe pagarse el día del evento antes de comenzar el servicio.
      </div>
      <? } ?>
      <div class="muted" style="margin-top:8px">
        <strong>Métodos de pago:</strong> Efectivo · ATH Móvil al <?= data.company.phone ?> · PayPal · Tarjetas de débito/crédito (≈3.99% por transacción)
      </div>

      <div class="muted" style="margin-top:12px">*Esta cotización es preliminar y está sujeta a confirmación de logística y distancia.</div>
    </div> <!-- cierre padding -->
  </div>   <!-- cierre card -->
</body>
</html>`;
}

/** Renderiza HTML -> PDF, lo guarda en /Cotizaciones y devuelve {file,url} */
function generateQuotePdf(row) {
  const folder = DriveApp.getFolderById(QUOTES_FOLDER_ID);
  const html = HtmlService.createTemplate(quoteTemplateHtml());
  html.data = mapRowForTemplate(row);

  const out = html.evaluate().setWidth(794).setHeight(1123); // A4 aprox.
  const blob = out.getAs('application/pdf')
                  .setName(`${row.nro || 'COT'} - ${row.nombre || 'Cliente'}.pdf`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { file, url: file.getUrl() };
}

/** (Opcional) Enviar email con PDF; adjunta términos solo la 1ª vez */
function sendQuoteEmail(row, pdfFile, isUpdate) {
  const depNum = Number(row.deposito || 0);
  const isInvoice = depNum > 0;
  const docWordTitle = isInvoice ? 'Factura de barra de café' : 'Cotización para servicios de barra de café';
  const docWordLower = isInvoice ? 'factura' : 'cotización';
  const subject = isUpdate
    ? `Actualización de tu ${docWordTitle} ${row.nro || ''}`
    : `${docWordTitle} ${row.nro || ''}`;
  if (!row.email) return;

  // Normaliza entradas: aceptamos un DriveFile o {file,url}
  let fileObj = null;
  let pdfUrl  = '';
  if (pdfFile) {
    if (typeof pdfFile.getBlob === 'function') {
      fileObj = pdfFile;
      try { pdfUrl = pdfFile.getUrl(); } catch(e) { pdfUrl = ''; }
    } else if (pdfFile.file) {
      fileObj = pdfFile.file;
      pdfUrl = pdfFile.url || '';
    }
  }
  // Si no queremos incluir el enlace de Drive, vaciamos pdfUrl para que no se muestre botón ni link
  if (!INCLUDE_DRIVE_LINK) {
    pdfUrl = '';
  }

  const stripeText = isInvoice
    ? ''
    : `\n\nPara reservar tu fecha, se requiere un depósito fijo de $${FIXED_DEPOSIT_AMOUNT.toFixed(2)}.\nPuedes realizar el pago aquí:\n${STRIPE_DEPOSIT_URL}\n\nUna vez recibido el depósito, tu fecha quedará separada oficialmente.`;

  const plain = `Hola ${row.nombre || ''},\n\nAdjuntamos tu ${isUpdate ? 'actualización de ' : ''}${docWordLower} ${row.nro || ''}.${stripeText}\n\nSi necesitas cambios, respóndenos por este medio.${pdfUrl ? '\n\nLink: ' + pdfUrl : ''}\n`;

  const openBtn = pdfUrl ? `
    <div style="margin-top:14px">
      <a href="${pdfUrl}"
         style="display:inline-block;background:#0ea5e9;color:#fff;padding:10px 14px;border-radius:8px;text-decoration:none;font-weight:600">
         Abrir ${docWordLower} (PDF)
      </a>
    </div>` : '';

  const depositBtn = isInvoice ? '' : `
    <div style="margin-top:16px;padding:16px;background:#f0faf7;border:1px solid #d1f0e8;border-radius:10px">
      <p style="margin:0 0 6px 0;font-weight:700;color:#0B3D2E">Reserva tu fecha</p>
      <p style="margin:0 0 12px 0;color:#374151;font-size:14px">
        Para separar oficialmente tu fecha se requiere un depósito fijo de
        <strong>$${FIXED_DEPOSIT_AMOUNT.toFixed(2)}</strong>.
        Una vez recibido, tu fecha queda reservada.
      </p>
      <a href="${STRIPE_DEPOSIT_URL}"
         style="display:inline-block;background:#0B3D2E;color:#fff;padding:12px 20px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px">
        Reservar fecha &mdash; Dep&oacute;sito de $${FIXED_DEPOSIT_AMOUNT.toFixed(2)}
      </a>
    </div>`;

  const html = `
  <div style="background:#f6f7f9;padding:24px 0;font-family:Inter,Segoe UI,Roboto,Arial,sans-serif">
    <div style="max-width:680px;margin:0 auto;background:#fff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">
      <div style="display:flex;gap:16px;align-items:center;padding:16px 20px;border-bottom:1px solid #eee">
        <img src="cid:logo" alt="${COMPANY.name}" width="44" height="44" style="border-radius:10px;display:block"/>
        <div style="line-height:1.3">
          <div style="font-weight:700">${COMPANY.name}</div>
          <div style="font-size:12px;color:#6b7280">
            Barra de café para eventos en Puerto Rico · Tel: ${COMPANY.phone} · ${COMPANY.email}
          </div>
        </div>
      </div>
      <div style="padding:16px 20px">
        <p style="margin:0 0 8px 0">Hola <strong>${escapeHtml(row.nombre || '')}</strong>,</p>
        <p style="margin:0 0 12px 0;color:#111827">Adjuntamos tu ${isUpdate ? 'actualización de ' : ''}${docWordLower} <strong>${escapeHtml(row.nro || '')}</strong>.</p>
        ${depositBtn}
        ${openBtn}
      </div>
    </div>
  </div>`.trim();

  // Helper para evitar adjuntar el mismo archivo dos veces (por nombre+tamaño)
  const options = {
    name: COMPANY.name,
    htmlBody: html,
    replyTo: COMPANY.email,
    inlineImages: { logo: getLogoBlob() },
    attachments: []
  };
  const pushUnique = (blob) => {
    if (!blob) return;
    try {
      // Calculamos una llave simple por nombre+tamaño
      const key = `${blob.getName()}|${(blob.getBytes() || []).length}`;
      const exists = options.attachments.some(b => `${b.getName()}|${(b.getBytes() || []).length}` === key);
      if (!exists) options.attachments.push(blob);
    } catch (e) {
      options.attachments.push(blob);
    }
  };

  // Adjunta solo si está permitido
  if (ATTACH_QUOTE_PDF && fileObj) {
    const blob = fileObj.getBlob().setName(`${row.nro || (isInvoice ? 'Factura' : 'Cotizacion')}.pdf`);
    pushUnique(blob);
  }
  if (!isUpdate && TERMS_PDF_ID) {
    try { pushUnique(DriveApp.getFileById(TERMS_PDF_ID).getBlob()); } catch(e) {}
  }

  GmailApp.sendEmail(row.email, subject, plain, options);
}

/**
 * Revisa la hoja "Leads" y procesa filas aprobadas sin PDF.
 * La puedes usar con un trigger "On edit" (del Spreadsheet) o con uno "Time-driven".
 */
function onLeadChange() {
  const lock = LockService.getScriptLock();
  // Si no logramos el lock, salimos para evitar ejecuciones concurrentes (emails duplicados).
  if (!lock.tryLock(30000)) return;

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim().toLowerCase());

    // Devuelve índice 1-based; si no existe, devuelve 0
    const idx = (name) => {
      const i = headers.indexOf(String(name).toLowerCase());
      return i >= 0 ? (i + 1) : 0;
    };

    const cAprobado = idx('aprobado');
    const cPrecio   = idx('precio');
    const cDeposito = idx('deposito');
    const cNro      = idx('nro');
    const cPdfUrl   = idx('pdfurl');
    const cSentAt   = idx('sent_at');

    // Requeridas mínimas
    if (!cAprobado || !cPrecio) return;

    // lector seguro (evita row[-1])
    const cell = (rowArr, col1) => (col1 > 0 ? rowArr[col1 - 1] : '');

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

    for (let i = 0; i < values.length; i++) {
      const row      = values[i];
      const aprobadoOk = normalizeYes_(cell(row, cAprobado));
      const precio   = String(cell(row, cPrecio) || '').trim();
      const nro      = String(cell(row, cNro) || '').trim();
      const pdfUrl   = String(cell(row, cPdfUrl) || '').trim();
      const sentAt   = String(cell(row, cSentAt) || '').trim();

      // Solo procesar si: aprobado="si", hay precio y NO hay marca previa.
      // Cualquier valor en pdfUrl (incl. "PROCESSING") cuenta como "ya en proceso".
      const already = !!pdfUrl || !!sentAt;
      if (!aprobadoOk || !precio || already) continue;

      // Marcar inmediatamente para evitar carreras de onEdit
      if (cPdfUrl) sh.getRange(i + 2, cPdfUrl).setValue('PROCESSING');

      // Construye objeto "lead" desde la fila
      const lead = {};
      headers.forEach((h, k) => { lead[h] = row[k]; });

      // Número de cotización (si no existe)
      if (!nro) {
        lead.nro = buildQuoteNumber();
        if (cNro) sh.getRange(i + 2, cNro).setValue(lead.nro);
      } else {
        lead.nro = nro;
      }

      // Genera PDF
      let pdfInfo = null;
      try {
        pdfInfo = generateQuotePdf(lead);
      } catch (e) {
        // Limpia la marca si falló, para reintento
        if (cPdfUrl) sh.getRange(i + 2, cPdfUrl).setValue('');
        throw e;
      }

      // Envía email
      let emailSent = false;
      try {
        sendQuoteEmail(lead, pdfInfo, false);
        emailSent = true;
      } catch (e) {
        // continúa; igualmente guardamos el link
      }

      // Escribe de forma explícita por columna (evita desalinear si no son contiguas)
      if (cPdfUrl) sh.getRange(i + 2, cPdfUrl).setValue(pdfInfo && pdfInfo.url ? pdfInfo.url : '');
      if (cSentAt && emailSent) sh.getRange(i + 2, cSentAt).setValue(new Date());
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// NOTE: Use an INSTALLABLE trigger pointing to onSheetEdit (do NOT keep a simple onEdit).
function onSheetEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAME) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim().toLowerCase());

    const cAprob    = headers.indexOf('aprobado') + 1;
    const cPrecio   = headers.indexOf('precio') + 1;
    const cDeposito = headers.indexOf('deposito') + 1;
    const cPdfUrl   = headers.indexOf('pdfurl') + 1;
    const cNro      = headers.indexOf('nro') + 1;
    const cSentAt   = headers.indexOf('sent_at') + 1;

    if (cAprob < 1 || cPrecio < 1) return;

    const row = e.range.getRow();
    if (row === 1) return;

    const aprobadoRaw = sh.getRange(row, cAprob).getValue();
    const aprobadoOk  = normalizeYes_(aprobadoRaw);

    const precio   = String(sh.getRange(row, cPrecio).getValue() || '').trim();
    const pdfUrl   = (cPdfUrl > 0) ? String(sh.getRange(row, cPdfUrl).getValue() || '').trim() : '';

    // ✅ MEJORA: solo cuenta como "tiene PDF" si hay url real (no vacío y no PROCESSING)
    const hasPdf = (cPdfUrl > 0) && pdfUrl && pdfUrl !== 'PROCESSING';

    // 1) Flujo inicial: aprobado + precio + aún sin PDF real -> genera y envía cotización
    const needsInitial = (aprobadoOk && precio !== '' && !hasPdf);
    if (needsInitial) {
      onLeadChange(); // este procesa la fila y genera/envía
      return;
    }

    // 2) Flujo de actualización: si ya hubo PDF real y se edita depósito -> recalcular y reenviar
    const editedCol = e.range.getColumn();
    const depositColumnEdited = (cDeposito > 0 && editedCol === cDeposito);
    const hasAprob = aprobadoOk;

    if (depositColumnEdited && hasPdf && hasAprob) {
      // Construye objeto 'lead' desde la fila completa
      const lastCol = sh.getLastColumn();
      const rowValues = sh.getRange(row, 1, 1, lastCol).getValues()[0];
      const lead = {};
      headers.forEach((h, k) => lead[h] = rowValues[k]);

      // Si no hay nro, asígnalo
      if (cNro > 0) {
        let nro = String(sh.getRange(row, cNro).getValue() || '').trim();
        if (!nro) {
          nro = buildQuoteNumber();
          sh.getRange(row, cNro).setValue(nro);
        }
        lead.nro = nro;
      }

      // Genera PDF actualizado y envía email de actualización
      const pdfInfo = generateQuotePdf(lead);
      let emailSent = false;
      try {
        sendQuoteEmail(lead, pdfInfo, true); // isUpdate = true
        emailSent = true;
      } catch (err) {
        console.warn('sendQuoteEmail update fail:', err);
      }

      if (cPdfUrl > 0) sh.getRange(row, cPdfUrl).setValue(pdfInfo && pdfInfo.url ? pdfInfo.url : '');
      if (cSentAt > 0 && emailSent) sh.getRange(row, cSentAt).setValue(new Date());
    }
  } catch (err) {
    console.warn('onEdit guard:', err);
  }
}