'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { fixture, inventory, write, run, ok } = require('./fixtures/localization-consumers/fixture.cjs');

const pagination = 'ApveeSpfxReactToolkitPaginationStrings';
const stage = 'ApveeSpfxReactToolkitStageStrings';
const paginationPath = 'node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js';
const stagePath = 'node_modules/@apvee/spfx-react-toolkit/lib/controls/stage/loc/{locale}.js';
const read = f => JSON.parse(fs.readFileSync(f.config, 'utf8'));

test('adds every inventory registration after existing entries and reruns without changing any bytes', t => {
  const f = fixture(t);
  ok(run(f.cli, f.project));
  assert.deepEqual(read(f).localizedResources, {
    WebPartStrings: 'lib/webparts/example/loc/{locale}.js',
    [pagination]: paginationPath,
    [stage]: stagePath
  });
  const before = fs.readFileSync(f.config);
  ok(run(f.cli, f.project));
  assert.deepEqual(fs.readFileSync(f.config), before);
  assert.equal(fs.existsSync(path.join(f.packageRoot, 'lib')), false);
});

test('adds an absent localizedResources map without changing other properties or order', t => {
  const f = fixture(t, { config: { version: '2.0', externals: {}, bundles: {} } });
  ok(run(f.cli, f.project));
  assert.deepEqual(Object.keys(read(f)), ['version', 'externals', 'bundles', 'localizedResources']);
  assert.deepEqual(read(f).localizedResources, { [pagination]: paginationPath, [stage]: stagePath });
});

test('preserves original property order including integer-like keys in consumer objects and maps', t => {
  const f = fixture(t, { config: '{"9":{"4":"four","1":"one"},"2":[{"6":true,"3":null}],"localizedResources":{"8":"custom/eight/{locale}.js","3":"custom/three/{locale}.js"}}' });
  ok(run(f.cli, f.project));
  assert.equal(fs.readFileSync(f.config, 'utf8'), '{"9":{"4":"four","1":"one"},"2":[{"6":true,"3":null}],"localizedResources":{"8":"custom/eight/{locale}.js","3":"custom/three/{locale}.js","ApveeSpfxReactToolkitPaginationStrings":"node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js","ApveeSpfxReactToolkitStageStrings":"node_modules/@apvee/spfx-react-toolkit/lib/controls/stage/loc/{locale}.js"}}');
});

test('preserves custom and unknown consumer mappings and accepts them in check mode', t => {
  const f = fixture(t, { config: { localizedResources: { [stage]: 'translations/stage/{locale}.js', ConsumerStrings: 'custom/{locale}.js' } } });
  const result = run(f.cli, f.project);
  ok(result);
  assert.match(result.stdout, /custom.*ApveeSpfxReactToolkitStageStrings/i);
  assert.deepEqual(read(f).localizedResources, {
    [stage]: 'translations/stage/{locale}.js', ConsumerStrings: 'custom/{locale}.js', [pagination]: paginationPath
  });
  const before = fs.readFileSync(f.config);
  ok(run(f.cli, f.project, ['--check']));
  assert.deepEqual(fs.readFileSync(f.config), before);
});

for (const [name, text, expectedIndent, expectedEol, expectedFinal] of [
  ['tabs and CRLF without a final newline', '{\r\n\t"version": "2.0"\r\n}', '\t', '\r\n', false],
  ['four spaces and LF with a final newline', '{\n    "version": "2.0"\n}\n', '    ', '\n', true],
  ['twelve spaces beyond JSON.stringify native indentation limit', '{\n            "version": "2.0"\n}', '            ', '\n', false],
  ['compact JSON without a final newline', '{"version":"2.0"}', '', '', false]
]) test(`preserves formatting style: ${name}`, t => {
  const f = fixture(t, { config: text });
  ok(run(f.cli, f.project));
  const actual = fs.readFileSync(f.config, 'utf8');
  const expected = expectedIndent === '            '
    ? '{\n            "version": "2.0",\n            "localizedResources": {\n                        "ApveeSpfxReactToolkitPaginationStrings": "node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js",\n                        "ApveeSpfxReactToolkitStageStrings": "node_modules/@apvee/spfx-react-toolkit/lib/controls/stage/loc/{locale}.js"\n            }\n}'
    : JSON.stringify({ version: '2.0', localizedResources: { [pagination]: paginationPath, [stage]: stagePath } }, undefined, expectedIndent)
      .replace(/\n/g, expectedEol) + (expectedFinal ? expectedEol : '');
  assert.equal(actual, expected);
});

test('check reports every missing ID and a repair command without changing config', t => {
  const f = fixture(t);
  const before = fs.readFileSync(f.config);
  const result = run(f.cli, f.project, ['--check']);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, new RegExp(pagination));
  assert.match(result.stderr, new RegExp(stage));
  assert.match(result.stderr, /spfx-react-toolkit configure-localization/);
  assert.deepEqual(fs.readFileSync(f.config), before);
});

test('list needs no project configuration or installed package and reports IDs, patterns and cultures', t => {
  const f = fixture(t);
  fs.rmSync(path.join(f.project, 'config'), { recursive: true });
  fs.rmSync(path.join(f.project, 'package.json'));
  fs.rmSync(path.join(f.packageRoot, 'package.json'));
  const result = run(f.cli, f.root, ['--list']);
  ok(result);
  for (const resource of inventory.resources) {
    assert.match(result.stdout, new RegExp(resource.id));
    assert.ok(result.stdout.includes(resource.path));
    assert.ok(result.stdout.includes('en-US'));
    assert.ok(result.stdout.includes('it-IT'));
  }
  assert.equal(fs.existsSync(f.config), false);
});

for (const [name, config] of [
  ['invalid strict JSON', '{"version":"2.0",}'],
  ['comments', '{/* unsupported */"version":"2.0"}'],
  ['root array', []],
  ['null map', { localizedResources: null }],
  ['array map', { localizedResources: [] }],
  ['numeric map', { localizedResources: 1 }],
  ['blank mapping', { localizedResources: { [stage]: '  ' } }],
  ['non-string known mapping', { localizedResources: { [stage]: {} } }],
  ['non-string unknown mapping', { localizedResources: { ConsumerStrings: false } }]
]) test(`rejects ${name} before adding any registration`, t => {
  const f = fixture(t, { config });
  const before = fs.readFileSync(f.config);
  const result = run(f.cli, f.project);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /config|localizedResources|JSON/i);
  assert.deepEqual(fs.readFileSync(f.config), before);
});

for (const args of [['--list', '--check'], ['--unknown'], ['--project'], ['--project', '--check'], ['unexpected'], ['--check', '--check'], ['--project', '.', '--project', '.']]) {
  test(`rejects invalid arguments ${args.join(' ')} without writes`, t => {
    const f = fixture(t);
    const before = fs.readFileSync(f.config);
    const result = run(f.cli, f.project, args);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /argument|option|exclusive|once|directory|Usage/i);
    assert.deepEqual(fs.readFileSync(f.config), before);
  });
}

test('rejects unknown commands without configuring a consumer', t => {
  const f = fixture(t);
  const before = fs.readFileSync(f.config);
  const result = spawnSync(process.execPath, [f.cli, 'configure-everything'], { cwd: f.project, encoding: 'utf8' });
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /command|Usage/i);
  assert.deepEqual(fs.readFileSync(f.config), before);
});

for (const [name, mutate] of [
  ['unsupported schema', data => { data.schemaVersion = 2; }],
  ['empty inventory', data => { data.resources = []; }],
  ['duplicate IDs', data => { data.resources.push(data.resources[0]); }],
  ['duplicate paths', data => { data.resources[1].path = data.resources[0].path; }],
  ['unsafe ID', data => { data.resources[0].id = '../unsafe'; }],
  ['traversal path', data => { data.resources[1].path = '../outside/{locale}.js'; }],
  ['absolute path', data => { data.resources[1].path = '/lib/controls/stage/loc/{locale}.js'; }],
  ['backslash path', data => { data.resources[1].path = 'lib\\controls\\stage\\loc\\{locale}.js'; }],
  ['missing locale placeholder', data => { data.resources[1].path = 'lib/controls/stage/loc/en-us.js'; }],
  ['unexpected metadata property', data => { data.resources[1].extra = true; }],
  ['missing en-US fallback', data => { data.resources[1].locales = ['it-IT']; }],
  ['duplicate cultures', data => { data.resources[1].locales.push('en-US'); }],
  ['invalid culture', data => { data.resources[1].locales.push('../en-US'); }]
]) test(`rejects metadata ${name} before writing a valid earlier entry`, t => {
  const f = fixture(t);
  const data = structuredClone(inventory);
  mutate(data);
  write(f.packageRoot, 'localization/resources.json', data);
  const before = fs.readFileSync(f.config);
  const result = run(f.cli, f.project);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /inventory|resource|locale|schema|metadata/i);
  assert.deepEqual(fs.readFileSync(f.config), before);
});

test('an upgrade appends newly introduced IDs while retaining custom and unknown entries', t => {
  const f = fixture(t, { inventory: { schemaVersion: 1, resources: [inventory.resources[0]] }, config: { localizedResources: { [pagination]: 'custom/pagination/{locale}.js', RetiredStrings: 'legacy/{locale}.js' } } });
  ok(run(f.cli, f.project));
  const before = fs.readFileSync(f.config);
  write(f.packageRoot, 'localization/resources.json', inventory);
  ok(run(f.cli, f.project));
  assert.deepEqual(read(f).localizedResources, { [pagination]: 'custom/pagination/{locale}.js', RetiredStrings: 'legacy/{locale}.js', [stage]: stagePath });
  assert.notDeepEqual(fs.readFileSync(f.config), before);
});

test('preserves a stale default as customization until the consumer removes that known mapping and reruns', t => {
  const stale = '../../old-layout/node_modules/@apvee/spfx-react-toolkit/lib/controls/pagination/loc/{locale}.js';
  const f = fixture(t, { config: { localizedResources: { [pagination]: stale, [stage]: 'custom/stage/{locale}.js', UnrelatedStrings: 'custom/other/{locale}.js' } } });
  ok(run(f.cli, f.project));
  assert.equal(read(f).localizedResources[pagination], stale);
  ok(run(f.cli, f.project, ['--check']));
  const repaired = read(f);
  delete repaired.localizedResources[pagination];
  write(f.project, 'config/config.json', repaired);
  ok(run(f.cli, f.project));
  assert.deepEqual(read(f).localizedResources, { [stage]: 'custom/stage/{locale}.js', UnrelatedStrings: 'custom/other/{locale}.js', [pagination]: paginationPath });
});

test('malformed consumer package JSON and a missing config fail without creating or rewriting files', t => {
  const f = fixture(t);
  const before = fs.readFileSync(f.config);
  write(f.project, 'package.json', '{bad');
  assert.notEqual(run(f.cli, f.project).status, 0);
  assert.deepEqual(fs.readFileSync(f.config), before);
  write(f.project, 'package.json', { name: 'consumer' });
  fs.rmSync(f.config);
  assert.notEqual(run(f.cli, f.project).status, 0);
  assert.equal(fs.existsSync(f.config), false);
});
