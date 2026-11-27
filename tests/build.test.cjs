const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const { validateConfig, resourceValues, renderManifest, escapeXml, validateCompilation, validateOutput, buildTarget } = require('../scripts/build.cjs');

const root = path.resolve(__dirname, '..');
const readConfig = target => JSON.parse(fs.readFileSync(path.join(root, 'config', `${target}.json`), 'utf8'));
const template = fs.readFileSync(path.join(root, 'manifests/manifest.template.xml'), 'utf8');

function hashes(directory) {
  const result = {};
  function visit(dir, prefix = '') {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const name = prefix + entry.name;
      if (entry.isDirectory()) visit(path.join(dir, entry.name), name + '/');
      else result[name] = crypto.createHash('sha256').update(fs.readFileSync(path.join(dir, entry.name))).digest('hex');
    }
  }
  visit(directory);
  return result;
}

function fixture(t) {
  const dist = path.join(root, 'dist');
  if (fs.existsSync(dist)) assert(!fs.lstatSync(dist).isSymbolicLink());
  fs.mkdirSync(dist, { recursive: true });
  const directory = fs.mkdtempSync(path.join(dist, '.test-'));
  for (const folder of ['config', 'public', 'manifests']) fs.cpSync(path.join(root, folder), path.join(directory, folder), { recursive: true });
  t.after(() => {
    assert.equal(path.dirname(directory), dist);
    assert(path.basename(directory).startsWith('.test-'));
    fs.rmSync(directory, { recursive: true, force: true });
  });
  return directory;
}

async function fakeCompile({ outputPath, entry, config }) {
  fs.writeFileSync(path.join(outputPath, `${entry}.js`), `/* ${config.clientId} ${config.authority} */`);
  return { assets: [{ name: `${entry}.js` }], chunks: [{ entry: true, initial: true }], warnings: [] };
}

test('Initial configurations preserve DEV identity and clearly separate PROD', () => {
  const dev = validateConfig(readConfig('dev'));
  const prod = validateConfig(readConfig('prod'));
  const legacy = fs.readFileSync(path.join(root, 'dev/manifest.xml'), 'utf8');
  const legacyAuth = fs.readFileSync(path.join(root, 'dev/authconfig.js'), 'utf8');
  assert(legacy.includes(`<Id>${dev.addinId}</Id>`));
  assert(legacy.includes(`<Version>${dev.manifestVersion}</Version>`));
  assert(legacyAuth.includes(dev.clientId));
  assert(legacyAuth.includes(dev.authority));
  assert.notEqual(prod.addinId, dev.addinId);
  assert.notEqual(prod.clientId, dev.clientId);
  assert(!fs.readFileSync(path.join(root, 'manifest.xml'), 'utf8').includes(prod.addinId));
  assert.equal(dev.configured, false);
  assert.equal(prod.configured, false);
});

test('URL normalization preserves subdirectories; AppDomain is origin only', () => {
  for (const baseUrl of ['https://example.invalid/nested/app', 'https://example.invalid/nested/app/']) {
    const values = resourceValues(validateConfig({ ...readConfig('dev'), baseUrl }));
    assert.equal(values.origin, 'https://example.invalid');
    assert.equal(values.javascriptRuntimeUrl, 'https://example.invalid/nested/app/graph-runtime.js');
    assert.equal(values.icon16Url, 'https://example.invalid/nested/app/icon-16.png');
  }
});

test('Invalid configuration fails without silently accepting placeholders or output paths', () => {
  const badValues = [
    { configured: 'false' }, { clientId: 'not-a-uuid' }, { addinId: '' },
    { baseUrl: 'http://example.invalid/' }, { baseUrl: 'https://example.invalid/?x=1' },
    { baseUrl: 'https://example.invalid/#fragment' }, { baseUrl: 'https://user:pass@example.invalid/' },
    { baseUrl: 'https://example.invalid/?' }, { baseUrl: 'https://example.invalid/\\path' },
    { authority: 'https://login.microsoftonline.com/common' }, { authority: 'https://example.com/22222222-2222-4222-8222-222222222222' },
    { manifestVersion: '1.0.0' }, { manifestVersion: '1.0.0.65536' },
    { displayName: ' ' }, { displayName: '\u0001bad' }, { displayName: '{{name}}' },
    { outputDir: '../dev' }, { configured: true },
  ];
  for (const override of badValues) assert.throws(() => validateConfig({ ...readConfig('dev'), ...override }));
  const missing = readConfig('dev');
  delete missing.authority;
  assert.throws(() => validateConfig(missing), /authority/);
  assert.throws(() => validateConfig({ ...readConfig('prod'), configured: true, baseUrl: 'https://example.com/' }), /test UUIDs/);
});

test('XML substitution escapes values, preserves the runtime contract and rejects incomplete templates', () => {
  const config = validateConfig({ ...readConfig('dev'), displayName: 'Test & <Example> "Name"', baseUrl: 'https://example.invalid/sub/' });
  const xml = renderManifest(template, config);
  assert(xml.includes('Test &amp; &lt;Example&gt; &quot;Name&quot;'));
  assert.equal((xml.match(/Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/g) || []).length, 2);
  assert(xml.includes('<Permissions>ReadWriteItem</Permissions>'));
  assert(xml.includes('<Override type="javascript" resid="JavaScriptRuntime.Url" />'));
  assert(xml.includes('<FunctionFile resid="WebRuntime.Url" />'));
  assert(xml.includes('<bt:Sets DefaultMinVersion="1.5">'));
  assert(xml.includes('<Sets DefaultMinVersion="1.1">'));
  assert(xml.includes('<MobileFormFactor>'));
  assert(!xml.includes('dima-09.github.io'));
  assert(!xml.includes('{{'));
  assert.throws(() => renderManifest(template + '{{unknown}}', config), /Unknown/);
  assert.throws(() => renderManifest(template.replace('{{addinId}}', 'fixed'), config), /Missing/);
  assert.throws(() => renderManifest(template + '{{broken_marker}}', config), /Unresolved/);
  assert.equal(escapeXml("<&\"'>"), '&lt;&amp;&quot;&apos;&gt;');
});

test('Manifest template changes only environment fields from the frozen DEV manifest', () => {
  const config = validateConfig({ ...readConfig('dev'), baseUrl: 'https://dima-09.github.io/outlook-graph-signature/dev/' });
  const rendered = renderManifest(template, config).replace(/outlook-graph-signature\/dev\/(icon-\d+\.png)/g, 'outlook-graph-signature/$1');
  const legacy = fs.readFileSync(path.join(root, 'dev/manifest.xml'), 'utf8').replace('  <!-- Replace only this GUID with your permanent add-in ID. -->\n', '');
  assert.equal(rendered, legacy);
});

test('Chunk validation rejects additional startup or lazy JavaScript', () => {
  assert.throws(() => validateCompilation({ assets: [{ name: 'signin.js' }, { name: 'vendor.js' }], chunks: [{ entry: true, initial: true }] }, 'signin'));
  assert.throws(() => validateCompilation({ assets: [{ name: 'signin.js' }], chunks: [{ entry: true, initial: true }, {}] }, 'signin'));
  assert.throws(() => validateCompilation({ assets: [{ name: 'signin.js' }], chunks: [{ entry: false, initial: true }] }, 'signin'));
});

test('Build replaces only its own environment, keeps licenses and detects mismatched resources', async t => {
  const directory = fixture(t);
  const log = [];
  const options = { root: directory, compile: async args => {
    const stats = await fakeCompile(args);
    fs.writeFileSync(path.join(args.outputPath, `${args.entry}.js.LICENSE.txt`), 'Test license');
    return stats;
  }, log: message => log.push(message) };
  const prod = await buildTarget('prod', options);
  const prodBefore = hashes(prod);
  const dev = await buildTarget('dev', options);
  assert.deepEqual(hashes(prod), prodBefore);
  const devBefore = hashes(dev);
  await buildTarget('prod', options);
  assert.deepEqual(hashes(dev), devBefore);
  await buildTarget('dev', options);
  assert.deepEqual(hashes(dev), devBefore);
  assert(log.some(message => message.includes('NOT CONFIGURED — DO NOT DEPLOY')));
  assert(fs.existsSync(path.join(dev, 'graph-runtime.js.LICENSE.txt')));
  assert.deepEqual(fs.readdirSync(path.join(directory, 'dist')).sort(), ['dev', 'prod']);
  const config = validateConfig(readConfig('dev'));
  const manifest = renderManifest(template, config);
  fs.writeFileSync(path.join(dev, '.well-known/microsoft-officeaddins-allowed.json'), JSON.stringify({ allowed: ['https://wrong.invalid/graph-runtime.js'] }));
  assert.throws(() => validateOutput(dev, config, manifest), /runtime URL/);
});

test('Compilation, validation and configuration failures preserve both previous builds', async t => {
  const directory = fixture(t);
  const options = { root: directory, compile: fakeCompile, log() {} };
  await buildTarget('dev', options);
  await buildTarget('prod', options);
  const before = hashes(path.join(directory, 'dist'));
  await assert.rejects(buildTarget('dev', { ...options, compile: async args => {
    if (args.entry === 'signin') throw new Error('Simulated compilation error');
    return fakeCompile(args);
  } }), /Simulated/);
  assert.deepEqual(hashes(path.join(directory, 'dist')), before);
  await assert.rejects(buildTarget('dev', { ...options, compile: async args => {
    const stats = await fakeCompile(args);
    fs.writeFileSync(path.join(args.outputPath, 'unexpected.js'), '');
    return stats;
  } }), /Unexpected JavaScript/);
  assert.deepEqual(hashes(path.join(directory, 'dist')), before);
  fs.writeFileSync(path.join(directory, 'config/dev.json'), JSON.stringify({ ...readConfig('dev'), clientId: '' }));
  await assert.rejects(buildTarget('dev', options), /clientId/);
  await assert.rejects(buildTarget('../../dev', options), /Unknown target/);
  assert.deepEqual(hashes(path.join(directory, 'dist')), before);
});

test('A failed output replacement restores the previous directory', async t => {
  const directory = fixture(t);
  const options = { root: directory, compile: fakeCompile, log() {} };
  await buildTarget('dev', options);
  await buildTarget('prod', options);
  const before = hashes(path.join(directory, 'dist'));
  const rename = fs.renameSync;
  let failed = false;
  t.mock.method(fs, 'renameSync', (source, destination) => {
    if (path.basename(source).startsWith('.tmp-dev-') && path.basename(destination) === 'dev') {
      failed = true;
      throw new Error('Simulated output replacement error');
    }
    return rename(source, destination);
  });
  await assert.rejects(buildTarget('dev', options), /Simulated output replacement/);
  assert(failed);
  assert.deepEqual(hashes(path.join(directory, 'dist')), before);
  assert.deepEqual(fs.readdirSync(path.join(directory, 'dist')).sort(), ['dev', 'prod']);
});

test('A same-target lock prevents concurrent replacement', async t => {
  const directory = fixture(t);
  fs.mkdirSync(path.join(directory, 'dist/.lock-dev'), { recursive: true });
  await assert.rejects(buildTarget('dev', { root: directory, compile: fakeCompile, log() {} }), /already being built/);
  assert(fs.existsSync(path.join(directory, 'dist/.lock-dev')));
});

test('Output junctions are rejected without touching their destination', async t => {
  const directory = fixture(t);
  const destination = path.join(directory, 'protected');
  fs.mkdirSync(destination);
  fs.writeFileSync(path.join(destination, 'keep.txt'), 'unchanged');
  fs.mkdirSync(path.join(directory, 'dist'));
  fs.symlinkSync(destination, path.join(directory, 'dist/dev'), process.platform === 'win32' ? 'junction' : 'dir');
  try {
    await assert.rejects(buildTarget('dev', { root: directory, compile: fakeCompile, log() {} }), /link|junction/);
    assert.equal(fs.readFileSync(path.join(destination, 'keep.txt'), 'utf8'), 'unchanged');
  } finally {
    fs.unlinkSync(path.join(directory, 'dist/dev'));
  }
});

test('CLI rejects unknown and extra arguments with a nonzero exit code', () => {
  for (const args of [['invalid'], ['dev', '--output', '../dev'], []]) {
    const result = spawnSync(process.execPath, [path.join(root, 'scripts/build.cjs'), ...args], { encoding: 'utf8' });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /Usage:/);
  }
});
