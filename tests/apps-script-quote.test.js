const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const source = fs.readFileSync('apps-script.gs', 'utf8');

function extractFunction(name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `No existe ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  throw new Error(`No se pudo extraer ${name}`);
}

function loadWorkflow(generateQuotePdf, sendQuoteEmail) {
  const context = { generateQuotePdf, sendQuoteEmail };
  vm.createContext(context);
  vm.runInContext(`${extractFunction('generateAndSendQuote')}; this.workflow = generateAndSendQuote;`, context);
  return context.workflow;
}

{
  const pdfInfo = { url: 'https://drive.test/quote', file: {} };
  const workflow = loadWorkflow(() => pdfInfo, () => ({ sent: true, recipient: 'client@example.com' }));
  const result = workflow({ nro: 'ML-TEST', email: 'client@example.com' }, false);
  assert.equal(result.sent, true);
  assert.equal(result.pdfInfo, pdfInfo);
  assert.equal(result.emailResult.sent, true);
}

{
  const workflow = loadWorkflow(() => ({ url: 'https://drive.test/quote' }), () => undefined);
  assert.throws(() => workflow({ nro: 'ML-TEST' }, false), /no confirmó el envío/);
}

{
  const gmailError = new Error('Gmail unavailable');
  const workflow = loadWorkflow(() => ({ url: 'https://drive.test/generated' }), () => { throw gmailError; });
  assert.throws(() => workflow({ nro: 'ML-TEST' }, false), gmailError);
  assert.equal(gmailError.generatedPdfUrl, 'https://drive.test/generated');
}

{
  const workflow = loadWorkflow(() => ({}), () => ({ sent: true }));
  assert.throws(() => workflow(null, false), /cotización válida/);
}

assert.match(source, /delivery\.sent === true\) sh\.getRange\([^\n]+sent_at|cSentAt && delivery\.sent === true/);

{
  const context = { assertAdminUser_: () => 'monselattepr@gmail.com' };
  vm.createContext(context);
  vm.runInContext(`${extractFunction('saveAdminQuoteDraft')}; this.saveDraft = saveAdminQuoteDraft;`, context);
  assert.throws(
    () => context.saveDraft({ sheetRow: 2, changes: { aprobado: 'si' } }),
    /Campos no permitidos: aprobado/
  );
}

{
  const previewSource = extractFunction('previewAdminQuote');
  assert.doesNotMatch(previewSource, /\.setValue|\.setValues|\.clearContent/);
  const context = { assertAdminUser_: () => 'monselattepr@gmail.com' };
  vm.createContext(context);
  vm.runInContext(`${previewSource}; this.preview = previewAdminQuote;`, context);
  assert.throws(() => context.preview({ sheetRow: 1 }), /fila de la cotización no es válida/);
}

{
  const sendSource = extractFunction('sendAdminQuote');
  assert.match(sendSource, /delivery\.sent !== true/);
  assert.match(sendSource, /generateAndSendQuote\(lead, wasSent\)/);
  assert.match(sendSource, /sentAtColumn\)\.setValue\(sentAt\)/);
  assert.ok(
    sendSource.indexOf('delivery.sent !== true') < sendSource.indexOf('sentAtColumn).setValue(sentAt)'),
    'sendAdminQuote debe confirmar sent === true antes de escribir sent_at'
  );
  assert.match(sendSource, /payload\.confirmResend !== true/);
}

console.log('Apps Script quote workflow tests passed.');
