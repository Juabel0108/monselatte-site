/* Browser regression: use an installed Playwright runtime through NODE_PATH.
 * node tests/wizard-hardening.test.js
 * All fetch calls are mocked; Apps Script and analytics requests are blocked.
 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');
const base = process.env.WIZARD_TEST_URL || 'http://127.0.0.1:8905/';
const screenshots = process.env.WIZARD_SCREENSHOTS;
const baseline = process.env.WIZARD_BASELINE;

(async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const context = await browser.newContext({ deviceScaleFactor: 1 });
    let remoteAttempts = 0;
    await context.route('**/script.google.com/**', route => { remoteAttempts++; return route.abort(); });
    await context.route('**/script.googleusercontent.com/**', route => { remoteAttempts++; return route.abort(); });
    await context.route('**/www.googletagmanager.com/**', route => route.fulfill({ body: '', contentType: 'application/javascript' }));
    await context.route('**/google-analytics.com/**', route => route.abort());
    await context.addInitScript(() => {
      window.__requests = [];
      window.fetch = (url, options) => new Promise((resolve, reject) => {
        window.__requests.push({ url, options });
        window.__resolveFetch = () => resolve({ type: 'opaque', status: 0, ok: false });
        window.__rejectFetch = () => reject(new TypeError('Simulated network failure'));
      });
    });
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    page.on('console', msg => { if (msg.type() === 'error') errors.push(msg.text()); });
    const countConversions = () => page.evaluate(() => window.dataLayer.filter(e => e[0] === 'event' && e[1] === 'generate_lead').length);
    const selectRadio = (name, value) => page.locator('label').filter({ has: page.locator(`input[name="${name}"][value="${value}"]`) }).click();
    const step = async n => {
      assert.equal(await page.locator('#wizardBar').getAttribute('aria-valuenow'), String(n));
      assert.match(await page.locator('#wizardBar').getAttribute('aria-valuetext'), new RegExp(`Paso ${n} de 5`));
      assert.equal(await page.locator('.wizard-step:visible').getAttribute('data-step'), String(n));
    };
    for (const [width, height] of [[320,568],[375,812],[390,844],[768,1024],[1024,768],[1440,900]]) {
      await page.setViewportSize({ width, height });
      await page.goto(`${base}?hardening=${width}&utm_source=test&utm_campaign=regression#reserva`);
      await page.evaluate(() => document.fonts.ready);
      await step(1);
      assert.equal(await page.locator('#wizardBar').getAttribute('role'), 'progressbar');
      assert.equal(await page.locator('#wizardFeedback').getAttribute('aria-live'), 'polite');
      // An invalid submit never starts a fetch or conversion.
      await page.evaluate(() => document.querySelector('#leadForm').requestSubmit());
      assert.equal(await countConversions(), 0);
      assert.equal(await page.evaluate(() => window.__requests.length), 0);
      await page.locator('#wizardBack').click(); // validation focuses contact first
      await page.locator('#wizardBack').click();
      await page.locator('#wizardBack').click();
      await page.locator('#wizardBack').click();
      await step(1);
      await selectRadio('tipo', 'Otro');
      const other = '<em id="summary-probe">Recepción</em>';
      const address = 'Calle <strong id="address-probe">123</strong>';
      await page.locator('#tipo_otro').fill(other);
      for (const [field, plus, minus, expected] of [['invitados','10','-10','60'],['horas_servicio','1','-1','3']]) {
        await page.locator(`[data-stepper="${field}"][data-delta="${plus}"]`).focus();
        await page.keyboard.press('Enter'); await step(1);
        assert.equal(await page.locator(`[name="${field}"]`).inputValue(), expected);
        await page.locator(`[data-stepper="${field}"][data-delta="${minus}"]`).focus();
        await page.keyboard.press('Enter'); await step(1);
      }
      assert.equal(await page.locator('#f-invitados').inputValue(), '50');
      await page.locator('#tipo_otro').press('Enter'); await step(2);
      await page.locator('#f-fecha').fill('2100-12-15');
      await page.locator('#f-hora-inicio').fill('10:00'); await page.locator('#f-hora-inicio').press('Tab');
      assert.equal(await page.locator('#f-hora-fin').inputValue(), '12:00');
      await page.locator('#f-hora-fin').fill('14:00'); await page.locator('#f-hora-fin').press('Tab');
      assert.equal(await page.locator('#f-horas-servicio').inputValue(), '4');
      await page.locator('#f-hora-inicio').fill('11:00'); await page.locator('#f-hora-inicio').press('Tab');
      assert.equal(await page.locator('#f-horas-servicio').inputValue(), '3');
      assert.match(await page.locator('#wizardFeedback').textContent(), /3 horas/);
      await page.locator('#multiDia').check();
      await page.locator('#wizardNext').click(); await step(2);
      assert.match(await page.locator('#errors').innerText(), /Selecciona la fecha de fin/);
      await page.locator('#f-fecha-fin').fill('2100-12-16');
      await page.locator('#multiDia').uncheck();
      assert.equal(await page.locator('#f-fecha-fin').inputValue(), '');
      await page.locator('#multiDia').check(); await page.locator('#f-fecha-fin').fill('2100-12-16');
      await page.locator('#wizardNext').click(); await step(3);
      await page.locator('#f-localidad').selectOption('San Juan'); await page.locator('#f-direccion').fill(address);
      await page.locator('#wizardNext').click(); await step(4);
      for (const value of ['Añadir bebidas frías','Añadir matcha','Espresso martini','Carajillo']) {
        await page.locator('label').filter({ has: page.locator(`input[value="${value}"]`) }).click();
      }
      await page.locator('#wizardNext').click(); await step(5);
      assert.equal(await page.locator('#wizardSummary em,#wizardSummary strong').count(), 0);
      assert.match(await page.locator('#wizardSummary').textContent(), /<em id="summary-probe">/);
      assert.match(await page.locator('#wizardSummary').textContent(), /11:00–14:00 \(3h\)/);
      const labels = ['evento','invitados','fecha','horario','lugar','menú'];
      for (const label of labels) {
        const box = await page.getByRole('button', { name: `Editar ${label}`, exact: true }).boundingBox();
        assert(box.width >= 44 && box.height >= 44);
      }
      await page.getByRole('button', { name: 'Editar horario', exact: true }).click(); await step(2);
      await page.locator('#wizardNext').click(); await page.locator('#wizardNext').click(); await page.locator('#wizardNext').click(); await step(5);
      await selectRadio('cliente_tipo', 'Empresa');
      assert.equal(await page.locator('#f-nombre').getAttribute('autocomplete'), 'organization');
      for (const value of ['Persona','Empresa']) {
        const box = await page.locator(`[name="cliente_tipo"][value="${value}"] + span`).boundingBox(); assert(box.height >= 44);
      }
      await page.locator('#f-nombre').fill('Empresa Prueba');
      await page.locator('#f-email').fill('test@example.com');
      await page.locator('#f-telefono').fill('(787) 555-0123');
      const message = width === 320 ? '=1+1' : width === 375 ? '   =1+1' : 'Prueba sin envío real.';
      await page.locator('#f-mensaje').fill(message);
      assert(!await page.evaluate(() => document.documentElement.scrollWidth > innerWidth));
      if (screenshots && [375,1440].includes(width)) {
        fs.mkdirSync(screenshots, { recursive: true });
        await page.locator('#wizardSummary').scrollIntoViewIfNeeded();
        await page.screenshot({ path: path.join(screenshots, `summary-${width}x${height}.png`) });
      }
      // Existing serializer on the same validated inputs must produce identical bytes.
      let expected;
      if (baseline) {
        const source = fs.readFileSync(path.join(baseline, 'main.js'), 'utf8');
        const serializer = source.slice(source.indexOf('async function saveToSheet('), source.indexOf('// Validaciones de formulario'));
        expected = await page.evaluate(async serializer => {
          let sent;
          const legacy = new Function('formData','fetch','MULTI_FIELDS','getUTM','SHEET_WEBAPP_URL', serializer + '\nreturn saveToSheet(formData);');
          await legacy(validateForm(document.querySelector('#leadForm')).data, (url, options) => { sent = {url,options}; return Promise.resolve(); }, MULTI_FIELDS, getUTM, SHEET_WEBAPP_URL);
          return sent;
        }, serializer);
      }
      const submit = page.locator('#leadForm button[type="submit"]');
      await submit.click();
      assert.equal(await page.locator('#leadForm').getAttribute('aria-busy'), 'true');
      assert(await submit.isDisabled()); assert.equal(await submit.textContent(), 'Enviando solicitud…');
      assert(await page.locator('#wizardSuccess').isHidden());
      assert.equal(await countConversions(), 1);
      await page.evaluate(() => {
        const form = document.querySelector('#leadForm');
        form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true, submitter: form.querySelector('[type="submit"]') }));
      });
      assert.equal(await page.evaluate(() => window.__requests.length), 1);
      assert.equal(await countConversions(), 1);
      const actual = await page.evaluate(() => window.__requests[0]);
      assert.equal(actual.options.method, 'POST'); assert.equal(actual.options.mode, 'no-cors');
      if (expected) assert.deepEqual(actual, expected);
      assert.deepEqual(JSON.parse(actual.options.body), {
        website: '', tipo: `Otro — ${other}`, invitados: '50', horas_servicio: '3',
        fecha: '2100-12-15', hora_inicio: '11:00', hora_fin: '14:00', multi_dia: 'si', fecha_fin: '2100-12-16',
        localidad: 'San Juan', direccion: address, cliente_tipo: 'Empresa', nombre: 'Empresa Prueba',
        email: 'test@example.com', telefono: '7875550123', mensaje: message,
        menu_frias: 'Añadir bebidas frías', menu_matcha: 'Añadir matcha', menu_licor: 'Espresso martini, Carajillo',
        utm_source: 'test', utm_medium: '', utm_campaign: 'regression', utm_content: '', utm_term: '', referrer: '', page_path: '/'
      });
      await page.evaluate(() => window.__rejectFetch());
      await page.waitForFunction(() => !document.querySelector('#leadForm [type="submit"]').disabled);
      assert.equal(await submit.textContent(), 'Enviar solicitud');
      assert.equal(await page.locator('#leadForm').getAttribute('aria-busy'), null);
      assert(await page.locator('#wizardSuccess').isHidden()); assert(await submit.evaluate(e => e === document.activeElement));
      assert.equal(await page.locator('#errors').innerText(), 'No pudimos enviar tu solicitud. Verifica tu conexión e inténtalo nuevamente.');
      assert.equal(await page.locator('#errors').getAttribute('aria-live'), 'polite');
      assert.equal(await page.locator('#f-mensaje').inputValue(), message);
      await submit.click(); assert.equal(await countConversions(), 2);
      assert.deepEqual(await page.evaluate(() => window.__requests[1]), actual, 'Retry must preserve the entire request');
      assert(await page.locator('#errors').isHidden());
      await page.evaluate(() => window.__resolveFetch());
      await page.waitForFunction(() => !document.querySelector('#wizardSuccess').classList.contains('hidden'));
      assert.equal(await page.locator('#leadForm').getAttribute('aria-busy'), null);
      assert.equal(await page.locator('#wizardSuccess h3').innerText(), 'SOLICITUD ENVIADA');
      assert.match(await page.locator('#wizardSuccess').innerText(), /Si se procesa correctamente/);
      assert(!await page.evaluate(() => document.documentElement.scrollWidth > innerWidth));
      if (screenshots && [375,1440].includes(width)) await page.screenshot({ path: path.join(screenshots, `sent-${width}x${height}.png`) });
      assert(await page.locator('#wizardSuccess h3').evaluate(e => e === document.activeElement));
      assert.equal(await page.evaluate(() => window.dataLayer.filter(e => e[0]==='event' && e[1]==='wizard_success_view').length), 1);
      console.log(`${width}x${height}: keyboard, steps, conditions, hours, safe summary, targets, simulated send/error/success, analytics and payload PASS`);
    }
    assert.equal(remoteAttempts, 0); assert.deepEqual(errors, []);
    console.log('Zero real submissions. No console errors.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
