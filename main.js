// Año dinámico en footer
const yearEl = document.getElementById('year');
if (yearEl) yearEl.textContent = new Date().getFullYear();

// Menú móvil
const menuBtn = document.getElementById('menuBtn');
const mobileMenu = document.getElementById('mobileMenu');
menuBtn?.addEventListener('click', () => mobileMenu.classList.toggle('hidden'));

// --- Config de contacto ---
const WA_NUMBER = '17876108953'; // Número real sin + ni espacios
const EMAIL_TO  = 'monselattepr@gmail.com';
// La URL del Apps Script vive en el <meta name="sheet-webapp-url"> de index.html.
// Al desplegar una nueva versión del Apps Script, actualiza SOLO ese meta tag —
// este archivo no debería tocarse.
const SHEET_WEBAPP_URL = document.querySelector('meta[name="sheet-webapp-url"]')?.content
  // ⚠️ FALLBACK DE EMERGENCIA: solo se usa si una página olvidó incluir el meta tag.
  // Puede quedar desactualizado respecto al deployment real — no confiar en él.
  || 'https://script.google.com/macros/s/AKfycbwuclOuyz199zYyU1jKpyAidYl_ef7FmYLikhjhZYxwI15agCY8gfokHRa0yvgGmN2A/exec';
if (!document.querySelector('meta[name="sheet-webapp-url"]')?.content) {
  console.warn('sheet-webapp-url: falta el meta tag en esta página; usando fallback hardcodeado (puede estar desactualizado).');
}

// GA4 seguro desde cualquier módulo (no-op si gtag no cargó)
function trackEvent(name, params) {
  try {
    if (typeof window.gtag !== 'function') return;
    window.gtag('event', name, Object.assign({ transport_type: 'beacon' }, params || {}));
  } catch (e) {
    console && console.warn && console.warn('gtag failed:', e);
  }
}

// --- Menú del wizard (EDITABLE POR EL DUEÑO) ---------------------------
// MENU_INCLUDED se muestra como "tu paquete ya lo incluye" (no seleccionable).
// MENU_CATALOG genera tarjetas checkbox de ADD-ONS con costo adicional.
// Si añades una categoría nueva, agrégala también a LEADS_HEADERS y
// sanitizeInput en apps-script.gs.
const MENU_INCLUDED = [
  'Espresso', 'Cortado', 'Latte', 'Americano', 'Cappuccino tradicional',
  'Chocolate caliente', 'Chai latte',
  'Sabores: vainilla, caramelo y temporada',
  'Leche regular + alternativas'
];
const MENU_CATALOG = [
  {
    field: 'menu_frias',
    label: 'Barra de bebidas frías',
    note: 'Incluye el menú completo servido sobre hielo: iced latte, iced americano, iced chai y más.',
    items: ['Añadir bebidas frías']
  },
  {
    field: 'menu_matcha',
    label: 'Matcha',
    note: 'Añade matcha preparado al momento, caliente o frío.',
    items: ['Añadir matcha']
  },
  {
    field: 'menu_licor',
    label: 'Cócteles de café',
    note: 'Add-on con costo adicional — licor incluido por Monselatte, para eventos 21+.',
    items: ['Espresso martini', 'Carajillo']
  }
];

// Campos del wizard que son checkboxes multi-valor (comparten name)
const MULTI_FIELDS = MENU_CATALOG.map(c => c.field);

// Header: sombra y fondo al hacer scroll
(function(){
  const header = document.querySelector('header');
  if (!header) return;
  window.addEventListener('scroll', () => {
    if (window.scrollY > 10) {
      header.classList.add('elevated','bg-brand-cream/95');
    } else {
      header.classList.remove('elevated','bg-brand-cream/95');
    }
  });
})();

// Nav link activo con IntersectionObserver
(function(){
  const navLinksContainer = document.getElementById('navLinks');
  if (!navLinksContainer) return;
  const links = navLinksContainer.querySelectorAll('a.nav-link');
  const linkMap = new Map();
  links.forEach(a => {
    const id = a.getAttribute('href')?.replace('#','');
    if (id) linkMap.set(id, a);
  });
  const sectionIds = ['sobre','paquetes','galeria','testimonios','contacto','reserva'];
  const sections = sectionIds.map(id => document.getElementById(id)).filter(Boolean);
  if (!sections.length) return;

  const io = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      const id = entry.target.id;
      const link = linkMap.get(id);
      if (!link) return;
      if (entry.isIntersecting) {
        links.forEach(el => el.classList.remove('nav-link-active','text-brand-green'));
        link.classList.add('nav-link-active','text-brand-green');
      }
    });
  }, { root: null, rootMargin: '0px 0px -65% 0px', threshold: 0.25 });

  sections.forEach(sec => io.observe(sec));
})();

// Reveal on scroll
(function(){
  const els = document.querySelectorAll('.reveal');
  if (!els.length) return;
  const revealIO = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add('reveal-visible');
          revealIO.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.12 }
  );
  els.forEach(el => revealIO.observe(el));
})();

// Guardar en Google Sheets (background)
async function saveToSheet(formData){
  try {
    const payload = Object.fromEntries(formData.entries());
    // Los checkboxes del menú comparten name: Object.fromEntries solo conserva
    // el último valor, así que los colapsamos aquí ("Latte, Cappuccino").
    MULTI_FIELDS.forEach(f => {
      const all = formData.getAll(f).map(v => String(v).trim()).filter(Boolean);
      payload[f] = all.join(', ');
    });
    Object.assign(payload, getUTM());

    // Merge "Otro" + detail into a single field for backend / Sheets / emails
    if ((payload.tipo || '').trim() === 'Otro') {
      const det = (payload.tipo_otro || '').trim();
      payload.tipo = det ? `Otro — ${det}` : 'Otro';
    }
    // No necesitamos enviar tipo_otro separado
    delete payload.tipo_otro;

    fetch(SHEET_WEBAPP_URL, {
      method: 'POST',
      mode: 'no-cors',
      body: JSON.stringify(payload)
    });
  } catch (err) {
    console.warn('No se pudo guardar en la hoja:', err);
  }
}

// Validaciones de formulario y anti-spam
function normalizePhone(value){
  return (value || '').replace(/[^0-9]/g, '');
}

/** ===== Accesibilidad de errores por campo ===== */
function getFieldEl(form, name) {
  return form?.querySelector(`[name="${name}"]`);
}
function clearAllFieldErrors(form){
  if (!form) return;
  form.querySelectorAll('.field-error').forEach(el => el.remove());
  form.querySelectorAll('[aria-invalid="true"]').forEach(el => {
    el.removeAttribute('aria-invalid');
    const describedBy = (el.getAttribute('aria-describedby') || '').split(/\s+/).filter(Boolean);
    const filtered = describedBy.filter(id => !id.startsWith('err-'));
    if (filtered.length) el.setAttribute('aria-describedby', filtered.join(' '));
    else el.removeAttribute('aria-describedby');
  });
}
function setFieldError(form, name, message){
  const el = getFieldEl(form, name);
  if (!el) return;
  el.setAttribute('aria-invalid', 'true');

  const id = `err-${name}`;
  let msgEl = form.querySelector(`#${id}`);
  if (!msgEl) {
    msgEl = document.createElement('p');
    msgEl.id = id;
    msgEl.className = 'field-error mt-1 text-xs text-red-700';
    // Radios/checkboxes en tarjetas: el input es sr-only dentro de un label,
    // así que el mensaje va al final del grupo, no pegado al input invisible.
    const group = el.closest(`[data-field-group="${name}"]`) || form.querySelector(`[data-field-group="${name}"]`);
    if (group) group.appendChild(msgEl);
    else if (el.insertAdjacentElement) el.insertAdjacentElement('afterend', msgEl);
    else if (el.parentElement) el.parentElement.appendChild(msgEl);
  }
  msgEl.textContent = message;

  const current = (el.getAttribute('aria-describedby') || '').split(/\s+/).filter(Boolean);
  if (!current.includes(id)) current.push(id);
  el.setAttribute('aria-describedby', current.join(' '));
}

function showErrors(messages){
  const box = document.getElementById('errors');
  if(!box) return;
  if(!messages.length){
    box.classList.add('hidden');
    box.innerHTML='';
    return;
  }
  box.classList.remove('hidden');
  box.innerHTML = '<ul class="list-disc pl-5">' + messages.map(m => `<li>${m}</li>`).join('') + '</ul>';
}
function todayISO(){ const d=new Date(); d.setHours(0,0,0,0); return d.toISOString().split('T')[0]; }

function computeServiceHours(horaInicio, horaFin){
  // Returns integer hours (1–12) or null if invalid
  if (!horaInicio || !horaFin) return null;
  const [sh, sm] = String(horaInicio).split(':').map(Number);
  const [eh, em] = String(horaFin).split(':').map(Number);
  if ([sh, sm, eh, em].some(n => Number.isNaN(n))) return null;

  const startMin = (sh * 60) + sm;
  const endMin   = (eh * 60) + em;
  const diffMin  = endMin - startMin;
  if (diffMin <= 0) return null;

  // Si escogen 1h30m, redondeamos hacia arriba a 2 horas
  const hrs = Math.ceil(diffMin / 60);
  if (hrs < 1 || hrs > 12) return null;
  return hrs;
}

// Acepta persona O empresa ("Café 787 & Co."): letras, números y puntuación
// básica, con al menos 2 letras (bloquea "12345"). Idéntica a la del backend.
const RE_NOMBRE = /^(?=(?:.*\p{L}){2})[\p{L}\p{N}&.,'’()\- ]{2,80}$/u;
const RE_EMAIL  = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

// Cada validador recibe el FormData y devuelve un mensaje de error o null.
const FIELD_VALIDATORS = {
  nombre(d){
    const v = (d.get('nombre')||'').trim();
    return RE_NOMBRE.test(v) ? null : 'Escribe tu nombre o el de tu empresa (2–80 caracteres, con letras).';
  },
  email(d){
    return RE_EMAIL.test((d.get('email')||'').trim()) ? null : 'Escribe un email válido.';
  },
  telefono(d){
    const tel = normalizePhone((d.get('telefono')||'').trim());
    return (tel.length >= 7 && tel.length <= 15) ? null : 'El teléfono debe tener entre 7 y 15 dígitos.';
  },
  fecha(d){
    const v = (d.get('fecha')||'').trim();
    if (!v) return 'Selecciona la fecha.';
    if (v < todayISO()) return 'Elige una fecha futura.';
    return null;
  },
  fecha_fin(d){
    const multi = (d.get('multi_dia')||'').trim();
    const fin   = (d.get('fecha_fin')||'').trim();
    if (multi && !fin) return 'Selecciona la fecha de fin de tu evento.';
    if (fin && fin < (d.get('fecha')||'').trim()) return 'La fecha de fin no puede ser anterior a la de inicio.';
    return null;
  },
  hora_inicio(d){
    return (d.get('hora_inicio')||'').trim() ? null : 'Selecciona la hora de inicio.';
  },
  hora_fin(d){
    const inicio = (d.get('hora_inicio')||'').trim();
    const fin    = (d.get('hora_fin')||'').trim();
    if (!fin) return 'Selecciona la hora de fin.';
    if (inicio && !computeServiceHours(inicio, fin)) {
      return 'La hora de fin debe ser posterior al inicio (máx. 12 horas).';
    }
    return null;
  },
  localidad(d){
    return (d.get('localidad')||'').trim() ? null : 'Selecciona un municipio.';
  },
  direccion(d){
    const v = (d.get('direccion')||'').trim();
    return (v && v.length >= 8) ? null : 'Escribe una dirección más detallada (≥ 8 caracteres).';
  },
  invitados(d){
    const n = parseInt(d.get('invitados')||'0', 10);
    return (n >= 1 && n <= 500) ? null : 'Debe estar entre 1 y 500.';
  },
  horas_servicio(d){
    const n = parseInt(d.get('horas_servicio')||'0', 10);
    return (n >= 1 && n <= 12) ? null : 'Selecciona entre 2 y 8 horas.';
  },
  tipo(d){
    return (d.get('tipo')||'').trim() ? null : 'Selecciona un tipo de evento.';
  },
  tipo_otro(d){
    const tipo = (d.get('tipo')||'').trim();
    if (tipo === 'Otro' && !(d.get('tipo_otro')||'').trim()) return 'Especifica el tipo de evento.';
    return null;
  },
};

// Qué campos valida cada paso del wizard
const STEP_FIELDS = {
  1: ['tipo', 'tipo_otro', 'invitados', 'horas_servicio'],
  2: ['fecha', 'fecha_fin', 'hora_inicio', 'hora_fin'],
  3: ['localidad', 'direccion'],
  4: [], // menú es opcional
  5: ['nombre', 'email', 'telefono']
};
const ALL_VALIDATED_FIELDS = Object.keys(FIELD_VALIDATORS);

// Corre los validadores indicados, pinta errores y enfoca el primero.
// Devuelve la lista de {name, message} (vacía si todo bien).
function runValidators(form, data, names){
  clearAllFieldErrors(form);
  const fieldErrs = [];
  names.forEach(name => {
    const fn = FIELD_VALIDATORS[name];
    if (!fn) return;
    const msg = fn(data);
    if (msg) fieldErrs.push({ name, message: msg });
  });

  fieldErrs.forEach(fe => setFieldError(form, fe.name, fe.message));
  showErrors(fieldErrs.map(fe => fe.message));

  if (fieldErrs.length){
    const first = getFieldEl(form, fieldErrs[0].name);
    if (first && typeof first.focus === 'function'){
      first.focus();
      if (typeof first.scrollIntoView === 'function'){
        first.scrollIntoView({behavior:'smooth', block:'center'});
      }
    }
  }
  return fieldErrs;
}

function validateForm(form){
  const data = new FormData(form);

  const hp = (data.get('website') || '').trim();
  if(hp){
    console.warn('Honeypot activado; bloqueo de envío.');
    return { ok:false, spam:true, data };
  }

  const errs = runValidators(form, data, ALL_VALIDATED_FIELDS);

  // Normalizaciones del payload (solo si vamos a enviar)
  const hora_inicio = (data.get('hora_inicio') || '').trim();
  const hora_fin    = (data.get('hora_fin') || '').trim();
  const computedHours = computeServiceHours(hora_inicio, hora_fin);
  if (computedHours) data.set('horas_servicio', String(computedHours));
  data.set('telefono', normalizePhone((data.get('telefono')||'').trim()));
  data.set('hora_inicio', hora_inicio);
  data.set('hora_fin', hora_fin);

  return { ok: errs.length === 0, spam:false, data, firstErrorField: errs[0]?.name };
}


// --- Helpers: Query string & UTM tracking ---
function getSearchParams() {
  try { return new URLSearchParams(window.location.search); } catch { return new URLSearchParams(); }
}
function getUTM() {
  const sp = getSearchParams();
  const utm = {
    utm_source: sp.get('utm_source') || '',
    utm_medium: sp.get('utm_medium') || '',
    utm_campaign: sp.get('utm_campaign') || '',
    utm_content: sp.get('utm_content') || '',
    utm_term: sp.get('utm_term') || '',
    referrer: document.referrer || '',
    page_path: location.pathname || '',
  };
  return utm;
}

// Fijar mínimo de fecha a hoy
document.addEventListener('DOMContentLoaded', () => {
  const fechaInput = document.querySelector('input[name="fecha"]');
  if (fechaInput) fechaInput.min = todayISO();
});

// Hora de fin auto-calculada: inicio + horas de servicio (del stepper).
// Si el usuario edita el fin manualmente (endTouched), dejamos de pisarla y
// pasamos a recalcular horas_servicio a partir de inicio+fin.
document.addEventListener('DOMContentLoaded', () => {
  const startEl = document.querySelector('input[name="hora_inicio"]');
  const endEl   = document.querySelector('input[name="hora_fin"]');
  const hrsEl   = document.querySelector('input[name="horas_servicio"]');
  if (!startEl || !endEl || !hrsEl) return;

  let endTouched = false;

  function autofillEnd(){
    if (endTouched) return;
    const start = startEl.value;
    const hrs = parseInt(hrsEl.value || '0', 10);
    if (!start || !(hrs >= 1 && hrs <= 12)) return;
    const [h, m] = start.split(':').map(Number);
    if ([h, m].some(Number.isNaN)) return;
    const total = ((h * 60 + m) + hrs * 60) % (24 * 60);
    endEl.value = `${String(Math.floor(total / 60)).padStart(2,'0')}:${String(total % 60).padStart(2,'0')}`;
  }

  startEl.addEventListener('change', autofillEnd);
  hrsEl.addEventListener('change', autofillEnd);
  endEl.addEventListener('change', () => {
    endTouched = true;
    const hrs = computeServiceHours(startEl.value, endEl.value);
    if (hrs) {
      hrsEl.value = String(hrs);
      hrsEl.dispatchEvent(new Event('input', { bubbles: true }));
    }
  });
});

// Eventos multi-día: el checkbox muestra/oculta la fecha de fin
document.addEventListener('DOMContentLoaded', () => {
  const multi = document.getElementById('multiDia');
  const wrap  = document.getElementById('fechaFinWrap');
  const fin   = document.querySelector('input[name="fecha_fin"]');
  const fecha = document.querySelector('input[name="fecha"]');
  if (!multi || !wrap || !fin) return;

  multi.addEventListener('change', () => {
    wrap.classList.toggle('hidden', !multi.checked);
    if (multi.checked) {
      fin.min = fecha?.value || todayISO();
      fin.focus();
    } else {
      fin.value = ''; // si desmarca, el evento vuelve a ser de un día
    }
  });
  fecha?.addEventListener('change', () => { fin.min = fecha.value || todayISO(); });
});

function buildMessage(formData){
  const tipo = (formData.get('tipo') || '').toString().trim();
  const tipoOtro = (formData.get('tipo_otro') || '').toString().trim();
  const tipoFinal = (tipo === 'Otro')
    ? (tipoOtro ? `Otro — ${tipoOtro}` : 'Otro')
    : tipo;

  return `Hola Monselatte, quiero cotizar una barra de café.\n\n` +
         `Nombre: ${formData.get('nombre')}\n` +
         `Email: ${formData.get('email')}\n` +
         `Teléfono: ${formData.get('telefono')}\n` +
         `Fecha: ${formData.get('fecha')}\n` +
         ((formData.get('fecha_fin') || '').toString().trim() ? `Fecha fin: ${formData.get('fecha_fin')}\n` : '') +
         `Hora de inicio: ${formData.get('hora_inicio')}\n` +
         `Horas de servicio: ${formData.get('horas_servicio')}\n` +
         `Hora de fin: ${formData.get('hora_fin')}\n` +
         `Localidad: ${formData.get('localidad')}\n` +
         `Dirección: ${formData.get('direccion')}\n` +
         `Tipo de evento: ${tipoFinal}\n` +
         `Invitados: ${formData.get('invitados')}\n` +
         buildExtrasBlock(formData) +
         `Mensaje: ${(formData.get('mensaje') || '—').toString().slice(0, 500)}`;
}

// Add-ons seleccionados para el mensaje.
function buildExtrasBlock(formData){
  const val = (n) => (formData.get(n) || '').toString().trim();
  const multi = (n) => formData.getAll(n).map(v => String(v).trim()).filter(Boolean).join(', ');

  const menu = [
    ['Bebidas frías (add-on)', multi('menu_frias')],
    ['Matcha (add-on)', multi('menu_matcha')],
    ['Cócteles de café (add-on)', multi('menu_licor')]
  ].filter(([,v]) => v);

  let out = '';
  if (menu.length)      out += '\nMenú deseado:\n' + menu.map(([k,v]) => `- ${k}: ${v}`).join('\n') + '\n';
  return out ? out + '\n' : '';
}

// Paso del wizard al que pertenece un campo (para saltar al primer error)
function stepForField(name){
  for (const [step, fields] of Object.entries(STEP_FIELDS)){
    if (fields.includes(name)) return parseInt(step, 10);
  }
  return null;
}

// Handler de envío: guarda en Sheets (el backend emailea la confirmación al
// cliente) y muestra el panel de éxito. WhatsApp queda como opción en el panel.
const form = document.getElementById('leadForm');
form?.addEventListener('submit', (e) => {
  e.preventDefault();
  const result = validateForm(form);
  if(!result.ok){
    const step = stepForField(result.firstErrorField);
    if (step && window.__wizardGoTo) window.__wizardGoTo(step, { keepErrors: true });
    return;
  }
  const data = result.data;
  saveToSheet(data);
  if (window.__wizardSuccess) window.__wizardSuccess(data);
});

// --- Lightbox de galería ---
(function initLightbox(){
  const imgs = Array.from(document.querySelectorAll('#galeria img'));
  if (!imgs.length) return;

  const modal    = document.getElementById('lightbox');
  const imgEl    = document.getElementById('lbImg');
  const caption  = document.getElementById('lbCaption');
  const btnClose = document.getElementById('lbClose');
  const btnPrev  = document.getElementById('lbPrev');
  const btnNext  = document.getElementById('lbNext');
  const backdrop = document.getElementById('lbBackdrop');

  let idx = 0;
  let touchX = null;

  function show(i){
    idx = (i + imgs.length) % imgs.length;
    const el = imgs[idx];
    imgEl.src = el.getAttribute('src');
    imgEl.alt = el.getAttribute('alt') || '';
    caption.textContent = el.getAttribute('alt') || '';
  }
  function open(i){
    show(i);
    backdrop.classList.remove('lb-backdrop-out');
    imgEl.classList.remove('lb-img-out');
    caption.classList.remove('lb-caption-out');

    backdrop.classList.add('lb-backdrop-in');
    imgEl.classList.add('lb-img-in');
    caption.classList.add('lb-caption-in');

    modal.classList.remove('hidden','lb-hidden');
    requestAnimationFrame(() => modal.classList.add('lb-visible'));
    document.body.classList.add('overflow-hidden');
  }
  function close(){
    backdrop.classList.remove('lb-backdrop-in');
    imgEl.classList.remove('lb-img-in');
    caption.classList.remove('lb-caption-in');

    backdrop.classList.add('lb-backdrop-out');
    imgEl.classList.add('lb-img-out');
    caption.classList.add('lb-caption-out');

    modal.classList.remove('lb-visible');

    setTimeout(() => {
      modal.classList.add('hidden','lb-hidden');
      backdrop.classList.remove('lb-backdrop-out');
      imgEl.classList.remove('lb-img-out');
      caption.classList.remove('lb-caption-out');
    }, 220);
    document.body.classList.remove('overflow-hidden');
  }
  function transitionTo(targetIndex){
    imgEl.classList.remove('lb-img-in');
    caption.classList.remove('lb-caption-in');
    imgEl.classList.add('lb-img-out');
    caption.classList.add('lb-caption-out');

    setTimeout(() => {
      imgEl.classList.remove('lb-img-out');
      caption.classList.remove('lb-caption-out');

      show(targetIndex);
      void imgEl.offsetWidth; // reflow
      imgEl.classList.add('lb-img-in');
      caption.classList.add('lb-caption-in');
    }, 200);
  }

  function prev(){ transitionTo(idx - 1); }
  function next(){ transitionTo(idx + 1); }

  imgs.forEach((el, i) => {
    el.classList.add('cursor-zoom-in');
    el.addEventListener('click', (e) => { e.preventDefault(); open(i); });
    el.setAttribute('tabindex', '0');
    el.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); open(i); }
    });
  });

  btnClose.addEventListener('click', close);
  backdrop.addEventListener('click', close);
  btnPrev.addEventListener('click', prev);
  btnNext.addEventListener('click', next);

  window.addEventListener('keydown', (e) => {
    if (modal.classList.contains('hidden')) return;
    if (e.key === 'Escape') close();
    else if (e.key === 'ArrowLeft') prev();
    else if (e.key === 'ArrowRight') next();
  });

  // Swipe en móvil
  imgEl.addEventListener('touchstart', (e) => { touchX = e.touches[0].clientX; }, {passive:true});
  imgEl.addEventListener('touchend', (e) => {
    if (touchX == null) return;
    const dx = e.changedTouches[0].clientX - touchX;
    if (Math.abs(dx) > 40) (dx > 0 ? prev() : next());
    touchX = null;
  });
})();

// --- FAQ acordeón (animación suave + accesibilidad) ---
(function initFAQ(){
  const items = Array.from(document.querySelectorAll('#faqList .faq-item'));
  if (!items.length) return;

  items.forEach(item => {
    const btn = item.querySelector('.faq-q');
    const panel = item.querySelector('.faq-a');

    btn.setAttribute('aria-expanded', 'false');
    btn.type = 'button';

    function open() {
      item.classList.add('is-open');
      btn.setAttribute('aria-expanded', 'true');
      panel.style.maxHeight = panel.scrollHeight + 'px';
      panel.style.opacity = '1';
    }
    function close() {
      item.classList.remove('is-open');
      btn.setAttribute('aria-expanded', 'false');
      panel.style.maxHeight = '0px';
      panel.style.opacity = '0';
    }
    function toggle() {
      const isOpen = item.classList.contains('is-open');
      // Cerrar otros
      items.forEach(i => { if (i !== item && i.classList.contains('is-open')) i.querySelector('.faq-q').click(); });
      isOpen ? close() : open();
    }

    // iniciar cerrado
    close();

    btn.addEventListener('click', toggle);
    btn.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
      if (e.key === 'ArrowDown') { items[(items.indexOf(item)+1)%items.length].querySelector('.faq-q').focus(); }
      if (e.key === 'ArrowUp')   { items[(items.indexOf(item)-1+items.length)%items.length].querySelector('.faq-q').focus(); }
    });
  });
})();

// --- Analytics (GA4) instrumentation ---------------------------------
(function () {
  if (typeof window.gtag !== 'function') return; // GA no cargó

  const send = trackEvent;

  window.addEventListener('DOMContentLoaded', function () {
    // 1) Nav clicks
    document.querySelectorAll('#navLinks a').forEach((a) => {
      a.addEventListener('click', function () {
        send('nav_click', {
          link: a.getAttribute('href') || '',
          text: (a.textContent || '').trim(),
        });
      });
    });

    // 2) CTAs hacia secciones
    document.querySelectorAll('a[href="#reserva"]').forEach((a) => {
      a.addEventListener('click', function () {
        send('cta_click', { cta: 'reserva', section: a.closest('section')?.id || 'header' });
      });
    });
    document.querySelectorAll('a[href="#paquetes"]').forEach((a) => {
      a.addEventListener('click', function () {
        send('cta_click', { cta: 'paquetes', section: a.closest('section')?.id || 'header' });
      });
    });

    // 3) Contacto (WhatsApp, email, teléfono)
    document.querySelectorAll(`a[href*="wa.me/${WA_NUMBER}"]`).forEach((a) => {
      a.addEventListener('click', function () {
        const location = a.closest('footer')
          ? 'footer'
          : a.classList.contains('fixed')
          ? 'floating'
          : 'body';
        send('contact', { method: 'whatsapp', location });
      });
    });
    document.querySelectorAll('a[href^="mailto:"]').forEach((a) => {
      a.addEventListener('click', function () {
        send('contact', { method: 'email' });
      });
    });
    document.querySelectorAll('a[href^="tel:"]').forEach((a) => {
      a.addEventListener('click', function () {
        send('contact', { method: 'phone' });
      });
    });

    // 4) Formulario (lead intent)
    const form = document.getElementById('leadForm');
    const emailBtn = document.getElementById('sendEmail');

    function readForm() {
      if (!form) return {};
      const fd = new FormData(form);
      const menuCount = MULTI_FIELDS.reduce((n, f) => n + fd.getAll(f).filter(Boolean).length, 0);
      return {
        form_channel: 'email', // se sobrescribe abajo
        municipio: (fd.get('localidad') || '').toString(),
        event_type: (fd.get('tipo') || '').toString(),
        guests: parseInt(fd.get('invitados') || '0', 10) || 0,
        event_date: (fd.get('fecha') || '').toString(),
        start_time: (fd.get('hora_inicio') || '').toString(),
        service_hours: parseInt(fd.get('horas_servicio') || '0', 10) || 0,
        end_time: (fd.get('hora_fin') || '').toString(),
        menu_items: menuCount,
        cliente_tipo: (fd.get('cliente_tipo') || '').toString(),
      };
    }

    // Envío por WhatsApp (submit del form)
    if (form) {
      form.addEventListener('submit', function (e) {
        const params = readForm();
        const ch =
          (e.submitter && e.submitter.dataset && e.submitter.dataset.channel) || 'whatsapp';
        params.form_channel = ch;
        send('generate_lead', params); // evento de conversión
        send('contact', { method: ch });
      });
    }

    // Click en "Enviar por Email"
    if (emailBtn && form) {
      emailBtn.addEventListener('click', function () {
        const params = readForm();
        params.form_channel = 'email';
        send('generate_lead', params);
        send('contact', { method: 'email' });
      });
    }
  });
})();
// Mostrar/ocultar input "Otro" en tipo de evento (tarjetas radio)
document.addEventListener('DOMContentLoaded', () => {
  const tipoRadios = document.querySelectorAll('input[name="tipo"]');
  const otroWrap = document.getElementById('tipoOtroWrap');
  if (tipoRadios.length && otroWrap) {
    tipoRadios.forEach(r => r.addEventListener('change', () => {
      otroWrap.classList.toggle('hidden', r.value !== 'Otro');
      if (r.value === 'Otro') document.getElementById('tipo_otro')?.focus();
    }));
  }
});

// ====== WIZARD DE COTIZACIÓN (5 pasos) ==================================
(function initWizard(){
  const form = document.getElementById('leadForm');
  if (!form) return;
  const steps = Array.from(form.querySelectorAll('.wizard-step'));
  if (!steps.length) return; // página sin wizard

  const TOTAL = steps.length;
  const bar     = document.getElementById('wizardBar');
  const label   = document.getElementById('wizardLabel');
  const count   = document.getElementById('wizardCount');
  const backBtn = document.getElementById('wizardBack');
  const nextBtn = document.getElementById('wizardNext');
  let current = 1;

  // --- Render del paso de menú: incluido (informativo) + add-ons (tarjetas) ---
  function renderMenu(){
    const host = document.getElementById('menuCatalog');
    if (!host) return;
    const included = `
      <div class="menu-included mt-4">
        <p class="menu-eyebrow">Incluido en tu servicio</p>
        <p class="font-serif text-xl text-brand-green mb-3">La experiencia Monselatte</p>
        <ul class="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-1 text-sm text-brand-text/80">
          ${MENU_INCLUDED.map(item => `<li class="flex gap-2"><span class="text-brand-green">✓</span><span>${item}</span></li>`).join('')}
        </ul>
      </div>
    `;
    host.innerHTML = included + MENU_CATALOG.map(cat => `
      <div class="menu-addon" data-field-group="${cat.field}">
        <div class="menu-addon__copy"><p class="menu-eyebrow">Extra opcional</p><p class="font-serif text-xl text-brand-green">${cat.label}</p>
        ${cat.note ? `<p class="mt-1 text-sm text-brand-text/65">${cat.note}</p>` : ''}</div>
        <div class="menu-addon__choices">
          ${cat.items.map(item => `
            <label class="cursor-pointer">
              <input type="checkbox" name="${cat.field}" value="${item}" class="peer sr-only" />
              <span class="menu-choice"><span>${item}</span><span class="menu-choice__mark" aria-hidden="true">+</span></span>
            </label>
          `).join('')}
        </div>
      </div>
    `).join('');
  }

  // --- Feedback dinámico del paso 1 (futuro slot del precio estimado) ---
  function updateFeedback(){
    const fb = document.getElementById('wizardFeedback');
    if (!fb) return;
    const inv = parseInt(form.querySelector('[name="invitados"]')?.value || '0', 10);
    const hrs = parseInt(form.querySelector('[name="horas_servicio"]')?.value || '0', 10);
    if (!(inv >= 1 && hrs >= 1)) { fb.hidden = true; return; }
    const baristas = inv > 120 ? 'dos baristas' : 'un barista profesional';
    fb.innerHTML = `☕ Para <strong>${inv} invitados</strong> durante <strong>${hrs} horas</strong>, solemos asignar <strong>${baristas}</strong> con estación completa.`;
    fb.hidden = false;
  }

  // --- Resumen del paso 5 ---
  function buildSummary(){
    const host = document.getElementById('wizardSummary');
    if (!host) return;
    const fd = new FormData(form);
    const val = (n) => (fd.get(n) || '').toString().trim();
    const multi = (n) => fd.getAll(n).map(v => String(v).trim()).filter(Boolean).join(', ');
    const tipo = val('tipo') === 'Otro' && val('tipo_otro') ? `Otro — ${val('tipo_otro')}` : val('tipo');
    const adds = MULTI_FIELDS.map(f => multi(f)).filter(Boolean).join(' · ');
    const fecha = val('fecha_fin') ? `${val('fecha')} → ${val('fecha_fin')}` : val('fecha');

    const rows = [
      ['Evento', tipo, 1],
      ['Invitados', val('invitados'), 1],
      ['Fecha', fecha, 2],
      ['Horario', val('hora_inicio') && val('hora_fin') ? `${val('hora_inicio')}–${val('hora_fin')} (${val('horas_servicio')}h)` : '', 2],
      ['Lugar', [val('localidad'), val('direccion')].filter(Boolean).join(' · '), 3],
      ['Menú', adds ? `Paquete básico + ${adds}` : 'Paquete básico', 4]
    ].filter(([,v]) => v);

    host.innerHTML = rows.map(([k, v, step]) => `
      <div class="flex items-baseline justify-between gap-3">
        <dt class="shrink-0 font-medium text-brand-text">${k}</dt>
        <dd class="text-right">${v}
          <button type="button" class="wizard-edit ml-1 text-brand-green underline text-xs" data-goto="${step}">editar</button>
        </dd>
      </div>
    `).join('');
  }

  // --- Navegación ---
  function goToStep(n, opts){
    n = Math.min(Math.max(1, n), TOTAL);
    steps.forEach(s => { s.hidden = (parseInt(s.dataset.step, 10) !== n); });
    current = n;

    const name = steps[n-1].dataset.stepName || '';
    if (label) label.textContent = `Paso ${n} de ${TOTAL} — ${name}`;
    if (count) count.textContent = `${n}/${TOTAL}`;
    if (bar)   bar.style.width = `${(n / TOTAL) * 100}%`;

    // style.display (no el atributo hidden): las clases inline-flex de Tailwind
    // tienen la misma especificidad que [hidden] y lo anulan por orden de hoja.
    if (backBtn) backBtn.style.display = (n === 1) ? 'none' : '';
    if (nextBtn) nextBtn.style.display = (n === TOTAL) ? 'none' : ''; // en el último paso mandan los botones de envío

    if (!(opts && opts.keepErrors)) { clearAllFieldErrors(form); showErrors([]); }
    if (n === TOTAL) buildSummary();

    // En el init NO robamos foco ni scrolleamos (la página acaba de cargar)
    if (!(opts && opts.silent)) {
      const heading = steps[n-1].querySelector('.wizard-heading');
      if (heading) heading.focus({ preventScroll: true });
      const anchor = document.getElementById('wizardTop') || steps[n-1];
      const rect = anchor.getBoundingClientRect();
      // Conserva la posición del usuario. Solo corrige si el progreso quedó
      // oculto por encima del header o completamente fuera de la pantalla.
      if (rect.top < 88 || rect.top > window.innerHeight - 100) {
        window.scrollTo({ top: window.scrollY + rect.top - 96, behavior: 'auto' });
      }
    }
  }
  // Los handlers de envío (fuera de este closure) la usan para saltar a errores
  window.__wizardGoTo = goToStep;

  // --- Panel de éxito tras el envío ---
  function showSuccess(formData){
    steps.forEach(s => { s.hidden = true; });
    if (backBtn) backBtn.style.display = 'none';
    if (nextBtn) nextBtn.style.display = 'none';
    const top = document.getElementById('wizardTop');
    if (top) top.style.display = 'none';
    showErrors([]);

    const panel = document.getElementById('wizardSuccess');
    if (!panel) return;
    const wa = document.getElementById('successWhatsApp');
    if (wa) wa.href = `https://wa.me/${WA_NUMBER}?text=${encodeURIComponent(buildMessage(formData))}`;
    panel.classList.remove('hidden');
    const heading = panel.querySelector('.wizard-heading');
    if (heading) heading.focus({ preventScroll: true });
    const rect = panel.getBoundingClientRect();
    if (rect.top < 88 || rect.top > window.innerHeight - 120) {
      window.scrollTo({ top: window.scrollY + rect.top - 96, behavior: 'auto' });
    }
    trackEvent('wizard_success_view', {});
  }
  window.__wizardSuccess = showSuccess;

  function tryAdvance(){
    const data = new FormData(form);
    const errs = runValidators(form, data, STEP_FIELDS[current] || []);
    if (errs.length) return;
    trackEvent('wizard_step_complete', { step: current, step_name: steps[current-1].dataset.stepName || '' });
    goToStep(current + 1);
    trackEvent('wizard_step_view', { step: current, step_name: steps[current-1].dataset.stepName || '' });
  }

  nextBtn?.addEventListener('click', tryAdvance);
  backBtn?.addEventListener('click', () => goToStep(current - 1));

  // Enter en pasos 1–4 avanza en vez de enviar
  form.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    if (e.target.tagName === 'TEXTAREA') return;
    if (current < TOTAL) { e.preventDefault(); tryAdvance(); }
  });

  // Links "editar" del resumen
  form.addEventListener('click', (e) => {
    const btn = e.target.closest('.wizard-edit');
    if (btn) goToStep(parseInt(btn.dataset.goto, 10));
  });

  // --- Steppers (− / +) ---
  form.addEventListener('click', (e) => {
    const btn = e.target.closest('.stepper-btn');
    if (!btn) return;
    const input = form.querySelector(`input[name="${btn.dataset.stepper}"]`);
    if (!input) return;
    const min = parseInt(input.min || '0', 10);
    const max = parseInt(input.max || '999', 10);
    const delta = parseInt(btn.dataset.delta || '1', 10);
    const cur = parseInt(input.value || '0', 10) || 0;
    input.value = String(Math.min(Math.max(cur + delta, min), max));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  });

  // Clamp al teclear en los inputs de los steppers + feedback vivo
  ['invitados', 'horas_servicio'].forEach(name => {
    const input = form.querySelector(`input[name="${name}"]`);
    input?.addEventListener('change', () => {
      const min = parseInt(input.min || '0', 10);
      const max = parseInt(input.max || '999', 10);
      const cur = parseInt(input.value || '0', 10);
      if (!Number.isNaN(cur)) input.value = String(Math.min(Math.max(cur, min), max));
      updateFeedback();
    });
    input?.addEventListener('input', updateFeedback);
  });

  // --- Toggle Persona/Empresa (cambia label, autocomplete y ejemplo) ---
  const nombreInput = document.getElementById('f-nombre');
  const nombreLabel = document.getElementById('nombre-label');
  form.querySelectorAll('input[name="cliente_tipo"]').forEach(r => {
    r.addEventListener('change', () => {
      const empresa = r.value === 'Empresa';
      if (nombreLabel) nombreLabel.textContent = empresa ? 'Nombre de la empresa*' : 'Nombre y Apellidos*';
      if (nombreInput) {
        nombreInput.setAttribute('autocomplete', empresa ? 'organization' : 'name');
        nombreInput.placeholder = empresa ? 'Jubilin Entertainment LLC' : '';
      }
    });
  });

  // --- Init ---
  renderMenu();
  updateFeedback();
  goToStep(1, { silent: true });

  // Funnel: primer paso visto cuando la sección entra al viewport
  const reserva = document.getElementById('reserva');
  if (reserva && 'IntersectionObserver' in window) {
    const io = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          trackEvent('wizard_step_view', { step: 1, step_name: steps[0].dataset.stepName || '' });
          io.disconnect();
        }
      });
    }, { threshold: 0.3 });
    io.observe(reserva);
  }
})();
