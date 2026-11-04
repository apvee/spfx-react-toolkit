const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const test = require('node:test');
const h = require('./fixtures/generated-tooling/helpers.cjs');
const app = 'apps/spfx-react-toolkit-test';
const npm = (root, args, env = {}) => spawnSync('npm', args, { cwd: root, encoding: 'utf8', timeout: 180000, env: { ...process.env, ...env } });

test('root/workspace build, serve, bundle, package and pack lifecycles never generate or check artifacts', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'generated-lifecycles-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const metadata = ['package.json', h.config.packageRoot + '/package.json', app + '/package.json'];
  for (const route of metadata) {
    fs.mkdirSync(path.dirname(path.join(root, route)), { recursive: true });
    fs.copyFileSync(path.join(h.repository, route), path.join(root, route));
  }
  const bin = path.join(root, 'node_modules/.bin'); fs.mkdirSync(bin, { recursive: true });
  // Persistent Gulp tasks and compiler are process doubles here. A separate test
  // below runs the real compiler; no server is left behind by this lifecycle probe.
  for (const name of ['tsc', 'gulp']) {
    const filename = path.join(bin, name);
    fs.writeFileSync(filename, '#!/bin/sh\nprintf "%s\\n" "$0 $*" >> "$LIFECYCLE_LOG"\n'); fs.chmodSync(filename, 0o755);
  }
  fs.mkdirSync(path.join(root, 'scripts/docs'), { recursive: true });
  for (const name of ['localization-resources', 'pack-library']) fs.writeFileSync(path.join(root, 'scripts/' + name + '.cjs'), "require('node:fs').appendFileSync(process.env.LIFECYCLE_LOG, process.argv.slice(1).join(' ') + '\\n');\n");
  for (const name of ['documentation', 'generated', 'executable-examples']) fs.writeFileSync(path.join(root, 'scripts/docs/' + name + '.cjs'), "throw new Error('Forbidden generation/check in lifecycle');\n");
  for (const route of metadata) {
    const scripts = JSON.parse(fs.readFileSync(path.join(root, route), 'utf8')).scripts;
    for (const [name, command] of Object.entries(scripts)) if (/^(?:(?:pre|post)?(?:build(?::.*)?|serve|bundle(?::.*)?|package:solution|pack(?::.*)?|publishOnly))$/.test(name)) {
      assert.doesNotMatch(command, /documentation\.cjs|generated\.cjs|executable-examples\.cjs|docs:(?:generate|check)|generated:(?:sync|check)|examples:(?:generate|check)/, route + ': ' + name);
    }
  }
  const log = path.join(root, 'lifecycle.log');
  const commands = [
    ['run', 'build'], ['run', 'serve'], ['run', 'bundle:ship'], ['run', 'package:solution'], ['run', 'pack:library'],
    ['run', 'build', '--workspace', '@apvee/spfx-react-toolkit'],
    ...['build', 'serve', 'bundle:ship', 'package:solution'].map(name => ['run', name, '--workspace', '@apvee/spfx-react-toolkit-test']),
    ['pack', '--workspace', '@apvee/spfx-react-toolkit', '--pack-destination', root],
    ['pack', '--workspace', '@apvee/spfx-react-toolkit-test', '--pack-destination', root],
    ['pack', '--pack-destination', root],
  ];
  for (const args of commands) {
    const result = npm(root, args, { LIFECYCLE_LOG: log });
    assert.equal(result.status, 0, args.join(' ') + '\n' + result.stdout + result.stderr);
    assert.doesNotMatch(result.stdout + result.stderr, /Forbidden generation/);
  }
  const calls = fs.readFileSync(log, 'utf8');
  assert.match(calls, /gulp serve/);
  assert.match(calls, /--check-inventory/);
  assert.match(calls, /--emit/);
});

test('real app compiler consumes copied compiled library and catches invalid/missing generated inputs without repair', { timeout: 300000 }, t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'generated-app-build-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  for (const route of ['README.md', 'LICENSE', 'AGENTS.md', 'docs', 'tests/fixtures', 'package.json', 'package-lock.json', h.config.packageRoot, app]) {
    fs.mkdirSync(path.dirname(path.join(root, route)), { recursive: true });
    fs.cpSync(path.join(h.repository, route), path.join(root, route), { recursive: true, filter: source => !/(?:^|\/)(?:node_modules|lib|dist|temp|release|sharepoint|\.docs)(?:\/|$)/.test(source) });
  }
  h.copyEngineInputs(root, ['scripts/localization-resources.cjs', 'scripts/pack-library.cjs']);
  fs.symlinkSync(path.join(h.repository, 'node_modules'), path.join(root, 'node_modules'), 'dir');
  // A root node_modules symlink contains the ORIGINAL workspace link. Override
  // that package at the app's nearest resolution boundary, preserving its peers.
  const scope = path.join(root, app, 'node_modules/@apvee'); fs.mkdirSync(scope, { recursive: true });
  fs.symlinkSync(path.join(root, h.config.packageRoot), path.join(scope, 'spfx-react-toolkit'), 'dir');
  const library = npm(root, ['run', 'build:library']);
  assert.equal(library.status, 0, library.stdout + library.stderr);
  const ts = require('typescript');
  const configFile = ts.readConfigFile(path.join(root, app, 'tsconfig.json'), ts.sys.readFile);
  const options = ts.parseJsonConfigFileContent(configFile.config, ts.sys, path.join(root, app)).options;
  const entry = path.join(root, app, 'src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/catalogue.ts');
  const resolved = ts.resolveModuleName('@apvee/spfx-react-toolkit/controls/web-part-shell', entry, options, ts.sys).resolvedModule;
  assert.ok(resolved && fs.realpathSync(resolved.resolvedFileName).startsWith(fs.realpathSync(path.join(root, h.config.packageRoot, 'lib')) + path.sep), 'app must resolve copied compiled lib, never original workspace');
  const compile = () => npm(root, ['run', 'typecheck', '--workspace', '@apvee/spfx-react-toolkit-test']);
  const valid = compile(); assert.equal(valid.status, 0, valid.stdout + valid.stderr);
  const original = fs.readFileSync(entry, 'utf8');
  const sentinel = path.join(path.dirname(entry), 'compiler-sentinel.tsx');
  fs.writeFileSync(sentinel, h.ownership + '\nexport const compilerSentinel: string = 42;\n');
  const before = fs.readFileSync(sentinel);
  const invalid = compile(); assert.notEqual(invalid.status, 0);
  assert.match(invalid.stdout + invalid.stderr, /TS2322/);
  assert.deepEqual(fs.readFileSync(sentinel), before);
  fs.unlinkSync(sentinel);
  fs.unlinkSync(entry);
  const missing = compile(); assert.notEqual(missing.status, 0);
  assert.match(missing.stdout + missing.stderr, /TS2307.*generated\/jsdoc-examples\/catalogue/);
  assert.equal(fs.existsSync(entry), false);
  fs.writeFileSync(entry, original);

  // This source passes the documentation checker, whose snippets intentionally
  // permit unused bindings, but fails the unchanged app noUnusedLocals contract.
  const source = path.join(root, h.config.sourceRoot, 'hooks/useStableCallback.ts');
  const prior = fs.readFileSync(source, 'utf8');
  const doc = [
    ' * @example Compiler boundary probe', ' * ```tsx', " * import * as React from 'react';",
    " * import { useState as unusedCompilerProbe } from 'react';",
    ' * export function Demo(): React.ReactElement { return <div>Compiler boundary</div>; }',
    " * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'Compiler boundary is visible.' };", ' * ```',
  ].join('\n');
  const commentEnd = prior.lastIndexOf(' */');
  fs.writeFileSync(source, prior.slice(0, commentEnd) + doc + '\n' + prior.slice(commentEnd));
  const { buildValidatedDocumentationContext } = require('../scripts/docs/documentation-context.cjs');
  const { renderExecutableArtifacts } = require('../scripts/docs/executable-examples.cjs');
  const context = buildValidatedDocumentationContext(root, h.config);
  const expected = renderExecutableArtifacts(root, context, h.config);
  const probeRoute = Object.keys(expected).find(route => expected[route].includes('unusedCompilerProbe'));
  assert.ok(probeRoute, 'checker-compatible source must be extracted');
  fs.writeFileSync(path.join(root, probeRoute), expected[probeRoute]);
  const unused = compile(); assert.notEqual(unused.status, 0);
  assert.match(unused.stdout + unused.stderr, /TS6133.*unusedCompilerProbe/);
  t.diagnostic((unused.stdout + unused.stderr).split('\n').filter(line => /TS6133.*unusedCompilerProbe/.test(line)).join('\n'));
  // Correct the source; do not weaken compiler options or patch generated bytes.
  fs.writeFileSync(source, fs.readFileSync(source, 'utf8').replace(" * import { useState as unusedCompilerProbe } from 'react';\n", ''));
  const repaired = renderExecutableArtifacts(root, buildValidatedDocumentationContext(root, h.config), h.config);
  fs.writeFileSync(path.join(root, probeRoute), repaired[probeRoute]);
  const green = compile(); assert.equal(green.status, 0, green.stdout + green.stderr);
  t.diagnostic('Real app compiler: empty catalogue accepted; TS2322/TS2307/TS6133 observed; source correction passes without compiler relaxation.');
});
