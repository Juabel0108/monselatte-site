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
const SHEET_WEBAPP_URL = (document.querySelector('meta[name="sheet-webapp-url"]')?.content) || 'https://script.google.com/macros/s/AKfycbzkfda9BR_FiaXCWIjdUSi2DH6evdxeF4OZufOkCmDZ1vDdzEZCxnPh2_xMYb4xhrvdtA/exec';

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
    const payload = Object.fromEntries(formData.entries()); Object.assign(payload, getUTM());
    fetch(SHEET_WEBAPP_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
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
function showErrors(messages){
  const box = document.getElementById('errors');
  if(!box) return;
  if(!messages.length){ box.classList.add('hidden'); box.innerHTML=''; return; }
  box.classList.remove('hidden');
  box.innerHTML = '<ul class="list-disc pl-5">' + messages.map(m => `<li>${m}</li>`).join('') + '</ul>';
}
function todayISO(){ const d=new Date(); d.setHours(0,0,0,0); return d.toISOString().split('T')[0]; }
function validateForm(form){
  const data = new FormData(form);
  const errors = [];
  const hp = (data.get('website') || '').trim();
  if(hp){ console.warn('Honeypot activado; bloqueo de envío.'); return { ok:false, spam:true, data }; }
  const nombre=(data.get('nombre')||'').trim();
  const email=(data.get('email')||'').trim();
  const telRaw=(data.get('telefono')||'').trim();
  const tel=normalizePhone(telRaw);
  const fecha=(data.get('fecha')||'').trim();
  const hora=(data.get('hora')||'').trim();
  const loc=(data.get('localidad')||'').trim();
  const direccion=(data.get('direccion')||'').trim();
  const invitados=parseInt(data.get('invitados')||'0',10);

  const reNombre=/^[A-Za-zÁÉÍÓÚÜÑáéíóúü' -]{2,60}$/;
  const reEmail=/^[^@\s]+@[^@\s]+\.[^@\s]+$/;

  if(!reNombre.test(nombre)) errors.push('Nombre: usa solo letras y espacios (2–60).');
  if(!reEmail.test(email)) errors.push('Email: formato inválido.');
  if(tel.length < 7 || tel.length > 15) errors.push('Teléfono: 7–15 dígitos.');
  if(!fecha) errors.push('Selecciona la fecha del evento.');
  if(!hora) errors.push('Selecciona el horario de inicio.');
  if(fecha && fecha < todayISO()) errors.push('La fecha no puede estar en el pasado.');
  if(!loc) errors.push('Selecciona el municipio.');
  if(!direccion || direccion.length < 8) errors.push('Dirección: escribe una dirección más detallada (mínimo 8 caracteres).');
  if(!(invitados >= 1 && invitados <= 500)) errors.push('Invitados: debe ser entre 1 y 500.');

  data.set('telefono', tel);
  showErrors(errors);
  return { ok: errors.length === 0, spam:false, data };
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
document.addEventListener('DOMContentLoaded', () => {
  // Preseleccionar paquete via ?paquete=Básico|Experiencia%20Monselatte|Premium
  const pkg = getSearchParams().get('paquete');
  if (pkg) {
    const sel = document.querySelector('select[name=\"paquete\"]');
    if (sel) {
      const options = Array.from(sel.options).map(o => o.value.toLowerCase());
      const idx = options.indexOf(pkg.toLowerCase());
      if (idx >= 0) sel.selectedIndex = idx;
    }
  }
});
// Fijar mínimo de fecha a hoy
document.addEventListener('DOMContentLoaded', () => {
  const fechaInput = document.querySelector('input[name="fecha"]');
  if (fechaInput) fechaInput.min = todayISO();
});

function buildMessage(formData){
  return `Hola Monselatte, quiero cotizar una barra de café.\n\n` +
         `Nombre: ${formData.get('nombre')}\n` +
         `Email: ${formData.get('email')}\n` +
         `Teléfono: ${formData.get('telefono')}\n` +
         `Fecha: ${formData.get('fecha')}\n` +
         `Hora de inicio: ${formData.get('hora')}\n` +
         `Localidad: ${formData.get('localidad')}\n` +
         `Dirección: ${formData.get('direccion')}\n` +
         `Tipo de evento: ${formData.get('tipo')}\n` +
         `Invitados: ${formData.get('invitados')}\n` +
         `Paquete: ${formData.get('paquete')}\n` +
         `Mensaje: ${formData.get('mensaje') || '—'}`;
}

// Handlers de envío
const form = document.getElementById('leadForm');
const formMsg = document.getElementById('formMsg');
form?.addEventListener('submit', (e) => {
  e.preventDefault();
  const result = validateForm(form);
  if(!result.ok) return;
  const data = result.data;
  const msg = buildMessage(data);
  saveToSheet(data);
  const url = `https://wa.me/${WA_NUMBER}?text=${encodeURIComponent(msg)}`;
  window.open(url, '_blank');
  formMsg?.classList.remove('hidden');
  document.getElementById('saveHint')?.classList.remove('hidden');
});

document.getElementById('sendEmail')?.addEventListener('click', () => {
  const result = validateForm(form);
  if(!result.ok) return;
  const data = result.data;
  const msg = buildMessage(data);
  saveToSheet(data);
  const subject = 'Solicitud de cotización — Monselatte';
  const mailto = `mailto:${EMAIL_TO}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(msg)}`;
  window.location.href = mailto;
  document.getElementById('saveHint')?.classList.remove('hidden');
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

  function send(name, params) {
    try {
      window.gtag('event', name, Object.assign({ transport_type: 'beacon' }, params || {}));
    } catch (e) {
      console && console.warn && console.warn('gtag failed:', e);
    }
  }

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
      return {
        form_channel: 'email', // se sobrescribe abajo
        municipio: (fd.get('localidad') || '').toString(),
        event_type: (fd.get('tipo') || '').toString(),
        package: (fd.get('paquete') || '').toString(),
        guests: parseInt(fd.get('invitados') || '0', 10) || 0,
        event_date: (fd.get('fecha') || '').toString(),
        event_time: (fd.get('hora') || '').toString(),
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