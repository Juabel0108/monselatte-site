const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

// Execute the real backend only with in-memory Sheets/email/cache doubles.
const context = vm.createContext({
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'local-test' }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  ContentService: {
    MimeType: { JSON: 'application/json' },
    createTextOutput: text => ({ text, setMimeType() { return this; } })
  },
  console
});
vm.runInContext(fs.readFileSync('apps-script.gs', 'utf8'), context);
const protect = context.safeSheetText_;
const dangerous = ['=1+1', '+1+1', '-1+1', '@SUM(A1:A2)', '   =1+1',
  '\t\r\n=1+1', '\u00a0+1+1', '\ufeff@SUM(A1:A2)', '\u0000=1+1', '= ejemplo'];
for (const value of dangerous) {
  assert.equal(protect(value), "'" + value);
  assert.equal(protect(value).slice(1), value, 'Preserve all original text after the Sheets marker');
  assert.equal(protect(protect(value)), protect(value), 'Do not double-escape');
}
for (const value of ['texto normal', 'correo@example.com', '2026-09-24', '10:00',
  '50', '', '   texto normal', "'=already literal", 0, 50, -10, 1.5, true, false,
  null, undefined, new Date('2026-09-24T12:00:00Z')]) {
  assert.equal(protect(value), value, 'Preserve legitimate value and type');
}

let writes, headers, spamRows, emails;
context.getOrCreateSheet = (_id, name, columns) => {
  if (name === 'Spam') return { appendRow: row => spamRows.push(row) };
  headers = Array.from(columns);
  return {
    getLastColumn: () => headers.length,
    getLastRow: () => 2,
    getRange: (row, col) => ({
      getValues: () => [headers],
      getDisplayValues: () => [headers.map(() => '')],
      setValue: value => { writes[headers[col - 1]] = value; }
    })
  };
};
context.isRateLimited = () => false;
context.sendAdminEmail = lead => emails.push(lead);
context.sendClientEmail = lead => emails.push(lead);
const fields = ['direccion','mensaje','tipo','paquete','menu_calientes','menu_frias',
  'menu_matcha','menu_licor','sabores','leches','utm_source','utm_medium','utm_campaign',
  'utm_term','utm_content','referrer','page_path'];
for (const value of [...dangerous, 'texto normal', 'correo@example.com']) {
  writes = {}; spamRows = []; emails = [];
  const data = {
    nombre: 'Cliente Prueba', email: 'correo@example.com', telefono: '7875550123',
    fecha: '2100-12-15', hora_inicio: '10:00', hora_fin: '12:00',
    horas_servicio: '2', invitados: '50', localidad: 'San Juan', website: ''
  };
  fields.forEach(field => { data[field] = value; });
  const result = context.doPost({ postData: { contents: JSON.stringify(data) } });
  assert.deepEqual(JSON.parse(result.text), { ok: true });
  const clean = context.sanitizeInput(data);
  for (const field of fields) {
    const original = Object.hasOwn(clean, field) ? clean[field] : data[field];
    assert.equal(writes[field], protect(original), field);
  }
  assert.equal(writes.nombre, data.nombre);
  assert.equal(writes.email, data.email);
  assert.equal(writes.localidad, data.localidad);
  assert.equal(Object.prototype.toString.call(writes.timestamp), '[object Date]');
  assert.equal(emails[0].mensaje, clean.mensaje, 'Sheets marker must not leak into email data');
  assert.equal(spamRows.length, 0);
  context.logSpam(value, value);
  assert.equal(spamRows[0][1], protect(value));
  assert.equal(spamRows[0][2], protect(value));
}
// Test the central writer independently of validation for every named lead field.
writes = {}; headers = ['timestamp','nombre','email','telefono','localidad','fecha','fecha_fin','hora_inicio'];
const sheet = context.getOrCreateSheet('', 'Leads', headers);
const external = Object.fromEntries(headers.map(field => [field, '  =1+1']));
context.appendRowByHeader_(sheet, external);
for (const field of headers) assert.equal(writes[field], "'  =1+1", field);
console.log('Sheets safety PASS: literal prefixes, whitespace, typed values, real doPost/write path, attribution, Spam; zero remote services.');
