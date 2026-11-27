const fs = require('node:fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const webpack = require('webpack');
const createWebpackConfig = require('../webpack.config.js');

const ROOT = path.resolve(__dirname, '..');
const TARGETS = ['dev', 'prod'];
const ENTRIES = ['graph-runtime', 'signin'];
const PUBLIC_FILES = ['graph-runtime.html', 'signin.html', 'commands.html', 'icon-16.png', 'icon-32.png', 'icon-80.png'];
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const TEST_IDS = ['11111111-1111-4111-8111-111111111111', '22222222-2222-4222-8222-222222222222'];
const WARNING = 'NOT CONFIGURED — DO NOT DEPLOY';

function assertTarget(target) {
  if (!TARGETS.includes(target)) throw new Error(`Unknown target: ${target}. Expected dev or prod.`);
}

function escapeXml(value) {
  for (const character of value) {
    const cp = character.codePointAt(0);
    if ((cp < 32 && ![9, 10, 13].includes(cp)) || (cp >= 0xd800 && cp <= 0xdfff) || cp === 0xfffe || cp === 0xffff) {
      throw new Error('Invalid XML character in configuration.');
    }
  }
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function httpsUrl(value, field) {
  let url;
  try { url = new URL(value); } catch { throw new Error(`${field} must be an absolute HTTPS URL.`); }
  if (url.protocol !== 'https:' || url.username || url.password || /[?#]/.test(value) || /[\s\\]/.test(value)) {
    throw new Error(`${field} must use HTTPS without credentials, query, fragment, whitespace or backslashes.`);
  }
  return url;
}

function validateConfig(input) {
  const fields = ['configured', 'baseUrl', 'clientId', 'authority', 'addinId', 'displayName', 'manifestVersion'];
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('Configuration must be a JSON object.');
  if (Object.keys(input).some(key => !fields.includes(key))) throw new Error('Unknown configuration field. Output paths are controlled by the builder.');
  if (typeof input.configured !== 'boolean') throw new Error('configured must be a boolean.');
  for (const field of fields.slice(1)) {
    if (typeof input[field] !== 'string' || !input[field].trim() || input[field] !== input[field].trim()) {
      throw new Error(`${field} must be a non-empty string without surrounding whitespace.`);
    }
    escapeXml(input[field]);
    if (input[field].includes('{{') || input[field].includes('}}')) throw new Error(`Unresolved marker in ${field}.`);
  }
  for (const field of ['clientId', 'addinId']) {
    if (!UUID.test(input[field])) throw new Error(`${field} must be a UUID.`);
  }
  if (!/^\d+\.\d+\.\d+\.\d+$/.test(input.manifestVersion) || input.manifestVersion.split('.').some(part => Number(part) > 65535)) {
    throw new Error('manifestVersion must have four numeric components in the range 0..65535.');
  }
  const base = httpsUrl(input.baseUrl, 'baseUrl');
  const authority = httpsUrl(input.authority, 'authority');
  const tenant = authority.pathname.replace(/^\//, '').replace(/\/$/, '');
  if (!UUID.test(tenant)) throw new Error('authority must contain a tenant UUID as its only path segment.');
  const isInvalidHost = url => /(^|\.)invalid$/i.test(url.hostname.replace(/\.$/, ''));
  if (authority.hostname !== 'login.microsoftonline.com' && !isInvalidHost(authority)) {
    throw new Error('authority must use login.microsoftonline.com, or an explicit .invalid test host.');
  }
  if (input.configured && (isInvalidHost(base) || isInvalidHost(authority) || [input.clientId, input.addinId, tenant].some(id => TEST_IDS.includes(id.toLowerCase())))) {
    throw new Error('configured=true cannot contain .invalid URLs or test UUIDs.');
  }
  base.pathname = base.pathname.replace(/\/+$/, '') + '/';
  return { ...input, baseUrl: base.href };
}

function resourceValues(config) {
  const url = file => new URL(file, config.baseUrl).href;
  return {
    addinId: config.addinId,
    displayName: config.displayName,
    manifestVersion: config.manifestVersion,
    origin: new URL(config.baseUrl).origin,
    icon16Url: url('icon-16.png'),
    icon32Url: url('icon-32.png'),
    icon80Url: url('icon-80.png'),
    commandsUrl: url('commands.html'),
    webRuntimeUrl: url('graph-runtime.html'),
    javascriptRuntimeUrl: url('graph-runtime.js'),
    signinUrl: url('signin.html'),
  };
}

function renderManifest(template, config) {
  const values = resourceValues(config);
  const used = new Set();
  const xml = template.replace(/\{\{([A-Za-z0-9]+)\}\}/g, (_, key) => {
    if (!Object.hasOwn(values, key)) throw new Error(`Unknown manifest marker: ${key}`);
    used.add(key);
    return escapeXml(values[key]);
  });
  if (xml.includes('{{') || xml.includes('}}')) throw new Error('Unresolved manifest marker.');
  for (const key of Object.keys(values)) if (!used.has(key)) throw new Error(`Missing manifest marker: ${key}`);
  return xml;
}

// All removable directories are direct children of dist. Reject links before
// writing, renaming or recursively removing any owned directory.
function assertNoLinks(directory) {
  const stat = fs.lstatSync(directory);
  if (stat.isSymbolicLink()) throw new Error(`Refusing symbolic link or junction: ${directory}`);
  if (stat.isDirectory()) {
    for (const child of fs.readdirSync(directory)) assertNoLinks(path.join(directory, child));
  }
}

function assertOwnedPath(root, directory) {
  const dist = path.resolve(root, 'dist');
  const resolved = path.resolve(directory);
  if (path.dirname(resolved) !== dist || !/^(dev|prod|\.lock-(dev|prod)|\.(tmp|backup)-(dev|prod)-[\w-]+)$/.test(path.basename(resolved))) {
    throw new Error('Refusing output path outside builder-owned dist directories.');
  }
  if (fs.existsSync(dist) && (fs.lstatSync(dist).isSymbolicLink() || !fs.lstatSync(dist).isDirectory())) {
    throw new Error('dist must be a regular directory, not a link.');
  }
  // lstat also detects dangling links, unlike existsSync.
  try { fs.lstatSync(resolved); } catch (error) { if (error.code === 'ENOENT') return; throw error; }
  if (fs.lstatSync(resolved).isSymbolicLink()) throw new Error(`Refusing symbolic link or junction: ${resolved}`);
  if (!fs.lstatSync(resolved).isDirectory()) throw new Error('Builder output must be a regular directory.');
  assertNoLinks(resolved);
}

function removeOwned(root, directory) {
  assertOwnedPath(root, directory);
  fs.rmSync(directory, { recursive: true, force: true });
}

function compileBundle({ root, outputPath, entry, config }) {
  return new Promise((resolve, reject) => {
    const compiler = webpack(createWebpackConfig({ root, outputPath, entry, config }));
    compiler.run((error, stats) => {
      compiler.close(closeError => {
        if (error || closeError) return reject(error || closeError);
        const result = stats.toJson({ all: false, errors: true, warnings: true, assets: true, chunks: true });
        if (stats.hasErrors()) return reject(new Error(result.errors.map(item => item.message).join('\n')));
        resolve(result);
      });
    });
  });
}

function validateCompilation(stats, entry) {
  const jsAssets = stats.assets.filter(asset => asset.name.endsWith('.js'));
  if (stats.chunks.length !== 1 || jsAssets.length !== 1 || jsAssets[0].name !== `${entry}.js` || !stats.chunks[0].entry || !stats.chunks[0].initial) {
    throw new Error(`${entry} must produce exactly one standalone JavaScript chunk.`);
  }
}

function listFiles(directory, prefix = '') {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap(entry => {
    const relative = prefix + entry.name;
    if (entry.isSymbolicLink()) throw new Error(`Unexpected output link: ${relative}`);
    return entry.isDirectory() ? listFiles(path.join(directory, entry.name), relative + '/') : [relative];
  }).sort();
}

function validateOutput(directory, config, expectedManifest) {
  const files = listFiles(directory);
  for (const file of [...PUBLIC_FILES, 'manifest.xml', '.well-known/microsoft-officeaddins-allowed.json', ...ENTRIES.map(entry => `${entry}.js`)]) {
    if (!files.includes(file)) throw new Error(`Missing output resource: ${file}`);
  }
  if (files.filter(file => file.endsWith('.js')).join('|') !== 'graph-runtime.js|signin.js') throw new Error('Unexpected JavaScript dependency in output.');
  const xml = fs.readFileSync(path.join(directory, 'manifest.xml'), 'utf8');
  if (xml !== expectedManifest) throw new Error('Manifest does not match the environment.');
  for (const entry of ENTRIES) {
    const js = fs.readFileSync(path.join(directory, `${entry}.js`), 'utf8');
    if (!js.includes(config.clientId) || !js.includes(config.authority) || /__AUTH_(CLIENT_ID|AUTHORITY)__/.test(js)) {
      throw new Error(`Incorrect authorization configuration in ${entry}.js`);
    }
    const html = fs.readFileSync(path.join(directory, `${entry}.html`), 'utf8');
    if (!html.includes(`src="./${entry}.js"`)) throw new Error(`Missing relative script reference in ${entry}.html`);
  }
  if (fs.readFileSync(path.join(directory, 'commands.html'), 'utf8').includes('commands.js')) throw new Error('Legacy commands.js must not be included.');
  const allowed = JSON.parse(fs.readFileSync(path.join(directory, '.well-known/microsoft-officeaddins-allowed.json'), 'utf8'));
  const runtimeUrl = resourceValues(config).javascriptRuntimeUrl;
  if (JSON.stringify(allowed) !== JSON.stringify({ allowed: [runtimeUrl] }) || !xml.includes(`id="JavaScriptRuntime.Url"\n            DefaultValue="${escapeXml(runtimeUrl)}"`)) {
    throw new Error('The allowed runtime URL must match the manifest exactly.');
  }
}

async function buildTarget(target, { root = ROOT, compile = compileBundle, log = console.log } = {}) {
  assertTarget(target);
  const config = validateConfig(JSON.parse(fs.readFileSync(path.join(root, 'config', `${target}.json`), 'utf8')));
  const template = fs.readFileSync(path.join(root, 'manifests', 'manifest.template.xml'), 'utf8');
  const manifest = renderManifest(template, config);
  const output = path.join(root, 'dist', target);
  const lock = path.join(root, 'dist', `.lock-${target}`);
  assertOwnedPath(root, output);
  assertOwnedPath(root, lock);
  fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
  try { fs.mkdirSync(lock); } catch (error) {
    if (error.code === 'EEXIST') throw new Error(`${target} is already being built (or has a stale lock at ${lock}).`);
    throw error;
  }
  let staging;
  let backup;
  try {
    log(`[${target}] ${config.configured ? 'CONFIGURED (deployment not verified)' : WARNING}`);
    staging = fs.mkdtempSync(path.join(root, 'dist', `.tmp-${target}-`));
    for (const entry of ENTRIES) {
      const stats = await compile({ root, outputPath: staging, entry, config });
      validateCompilation(stats, entry);
      for (const warning of stats.warnings || []) log(`[${target}/${entry}] WARNING: ${warning.message || warning}`);
    }
    for (const file of PUBLIC_FILES) fs.copyFileSync(path.join(root, 'public', file), path.join(staging, file));
    fs.writeFileSync(path.join(staging, 'manifest.xml'), manifest);
    fs.mkdirSync(path.join(staging, '.well-known'));
    fs.writeFileSync(path.join(staging, '.well-known/microsoft-officeaddins-allowed.json'), JSON.stringify({ allowed: [resourceValues(config).javascriptRuntimeUrl] }, null, 2) + '\n');
    validateOutput(staging, config, manifest);
    assertOwnedPath(root, staging);
    assertOwnedPath(root, output);
    if (fs.existsSync(output)) {
      backup = path.join(root, 'dist', `.backup-${target}-${randomUUID()}`);
      assertOwnedPath(root, backup);
      fs.renameSync(output, backup);
    }
    try { fs.renameSync(staging, output); } catch (error) {
      if (backup) fs.renameSync(backup, output);
      throw error;
    }
    if (backup) removeOwned(root, backup);
    log(`[${target}] Complete: ${output}${config.configured ? '' : ` — ${WARNING}`}`);
    return output;
  } finally {
    if (staging) removeOwned(root, staging);
    removeOwned(root, lock);
  }
}

async function main(args) {
  if (args.length !== 1 || ![...TARGETS, 'all'].includes(args[0])) throw new Error('Usage: node scripts/build.cjs <dev|prod|all>');
  for (const target of args[0] === 'all' ? TARGETS : args) await buildTarget(target);
}

module.exports = { validateConfig, resourceValues, renderManifest, escapeXml, validateCompilation, validateOutput, buildTarget, main };

if (require.main === module) {
  main(process.argv.slice(2)).catch(error => {
    console.error(`Build failed: ${error.message}`);
    process.exitCode = 1;
  });
}
