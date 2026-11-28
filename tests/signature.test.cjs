const test = require('node:test');
const assert = require('node:assert/strict');
const { createHash } = require('node:crypto');

test('Signature output is unchanged for fictional profiles and edge cases', async () => {
  const { buildSignature } = await import('../src/signature-template.js');
  const base = { displayName: 'Test User', jobTitle: 'Test Engineer', mail: 'user@example.invalid' };
  const cases = [
    {},
    base,
    { ...base, department: 'BDO Qazaqstan', businessPhones: ['123'], mobilePhone: '+0 000 000 00 00' },
    { ...base, department: '  bdo   outsourcing qazaqstan  ', businessPhones: ['456'] },
    { ...base, department: 'Unknown Test Company', businessPhones: [], mobilePhone: null },
    { displayName: '  Test <User> & "Example"  ', jobTitle: "R&D's", department: '<Company>', businessPhones: ['<123>'], mail: 'not <an> email' },
    { displayName: 'Test User', mail: '', userPrincipalName: 'ignored@example.invalid', department: ' ' },
    { displayName: 'Test User', mail: ' test+label@example.invalid ', department: null, businessPhones: [null] },
    { displayName: 123, jobTitle: false, department: {}, mail: 12, mobilePhone: [] },
  ];
  // SHA-256 of the complete original HTML for each fictional profile.
  const expected = [
    '94b329f36d56a318f4551c679e18fd7e83b039e1ead757a3a163a9f611660933',
    '289abca19dca760a0d33af18dcc4d116847ae7fd0cabdb6dbfbf2f90033c4138',
    '63b8a461ef8184ad6573771fce1bec382ed9f23ef350a6aac706ffc9b4b7cbb4',
    '003297fb285b11e627818b152d522bd2bf5e06b93719f4a8bfa1b326aa118ecf',
    'd59d81923ea2b7dcdaa73108edff44233ef2debf143abc06c3943c9861b1e435',
    '4df628b041efd435ccd5db9506d62c0f088401bc77bf39f3d39fe42f2d15af1e',
    '159faaa3e9cf7821c301ea41ec5d87ed3becdf544df766ec022ff7f1adb9a96e',
    '916f669af8840553c7feaf6641580957d3046b5022d8c3a005dbf14cc2e1efa4',
    '94b329f36d56a318f4551c679e18fd7e83b039e1ead757a3a163a9f611660933',
  ];
  assert.deepEqual(cases.map(p => createHash('sha256').update(buildSignature(p)).digest('hex')), expected);
  assert(buildSignature(cases[2]).includes('Direct: +7 727 331 31 34 ext. 123'));
  assert(buildSignature(cases[3]).includes('Direct: +7 727 331 31 37 ext. 456'));
  assert(buildSignature(cases[4]).includes('Tel: +7 727 331 31 34'));
  assert(buildSignature(cases[5]).includes('&lt;USER&gt; &amp; &quot;EXAMPLE&quot;'));
  assert(!buildSignature(cases[6]).includes('ignored@example.invalid'));
  assert(!buildSignature(cases[6]).includes('Direct:'));
  assert(!buildSignature(cases[5]).includes('mailto:'));
});
