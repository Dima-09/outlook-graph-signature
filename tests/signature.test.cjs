const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
function read(file) { return fs.readFileSync(path.join(root, file), 'utf8'); }
function generator(file) {
  const context = vm.createContext({});
  vm.runInContext(read(file).replace('export function buildSignature', 'function buildSignature'), context);
  return context.buildSignature;
}

test('Runtime, authentication, UI and public resources retain the frozen implementation', () => {
  for (const name of ['graph-auth.js', 'graph-runtime-source.js', 'signature-template.js', 'signin-source.js']) {
    assert.equal(read(`src/${name}`), read(`dev/${name}`), name);
  }
  for (const name of ['graph-runtime.html', 'signin.html']) assert.equal(read(`public/${name}`), read(`dev/${name}`));
  for (const name of ['icon-16.png', 'icon-32.png', 'icon-80.png']) {
    assert.deepEqual(fs.readFileSync(path.join(root, 'public', name)), fs.readFileSync(path.join(root, name)));
  }
  assert.equal(read('public/commands.html'), read('dev/commands.html').replace(/    <script src="\.\/commands\.js"><\/script>\r?\n\r?\n/, ''));
});

test('Signature output is unchanged for fictional profiles and edge cases', () => {
  const oldSignature = generator('dev/signature-template.js');
  const newSignature = generator('src/signature-template.js');
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
  for (const profile of cases) assert.equal(newSignature(profile), oldSignature(profile));
  assert(newSignature(cases[2]).includes('Direct: +7 727 331 31 34 ext. 123'));
  assert(newSignature(cases[3]).includes('Direct: +7 727 331 31 37 ext. 456'));
  assert(newSignature(cases[4]).includes('Tel: +7 727 331 31 34'));
  assert(newSignature(cases[5]).includes('&lt;USER&gt; &amp; &quot;EXAMPLE&quot;'));
  assert(!newSignature(cases[6]).includes('ignored@example.invalid'));
  assert(!newSignature(cases[6]).includes('Direct:'));
  assert(!newSignature(cases[5]).includes('mailto:'));
});
