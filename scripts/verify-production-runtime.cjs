const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const { createRequire } = require('node:module');
const Module = require('node:module');
const webpack = require('webpack');
const { buildConfig, statsOptions, collectEmittedModules, renderConsumer, readRuntimeToolchain, assertToolkitModuleOrigins } = require('./bundle-contract.internal.cjs');
const root = path.resolve(__dirname, '..');
const base = 'https://tenant.test/sites/standalone';
const toolkit = '@apvee/spfx-react-toolkit';
const scenarios = ['context', 'list', 'search', 'batch', 'providers', 'styles'];
const variants = ['root', 'domain', 'leaf'];
const leaf = { context: 'spfx-pnp-context.service', list: 'spfx-pnp-list.service', search: 'spfx-pnp-search.service', batch: 'spfx-pnp.service' };
const hash = data => crypto.createHash('sha256').update(data).digest('hex');

async function compileProbe({ consumerRoot, outputDir, scenario, variant }) {
  if (!scenarios.includes(scenario) || !['root', 'domain', 'leaf'].includes(variant)) throw new Error('Unknown production runtime probe');
  const ui = scenario === 'providers' || scenario === 'styles';
  const fixture = ui ? `${scenario}.tsx` : `pnp-${scenario}.ts`;
  const specifier = toolkit + (variant === 'root' ? '' : variant === 'domain' ? '/services' : '/lib/services/' + leaf[scenario]);
  const imports = ui ? scenario === 'providers' ? {
    PROVIDER: toolkit + (variant === 'root' ? '' : variant === 'domain' ? '/core' : '/lib/core/provider-webpart'),
    CONTEXT: toolkit + (variant === 'root' ? '/hooks' : ''),
    PROPERTIES: toolkit + (variant === 'root' ? '/lib/hooks/useSPFxProperties' : '/hooks'),
  } : {
    HOOK: toolkit + (variant === 'root' ? '' : variant === 'domain' ? '/styles' : '/lib/hooks/useSx'),
    OTHER_HOOK: toolkit + (variant === 'root' ? '/styles/useSx' : ''),
    DESCRIPTORS: toolkit + (variant === 'root' ? '/styles' : '/lib/helpers/styles'),
  } : {TOOLKIT:specifier};
  const source = renderConsumer(fs.readFileSync(path.join(root, 'tests/fixtures/tree-shaking/runtime', fixture), 'utf8'), { imports });
  const entry = path.join(consumerRoot, '.production-runtime', `${scenario}-${variant}${ui ? '.tsx' : '.ts'}`);
  fs.mkdirSync(path.dirname(entry), { recursive: true });
  fs.writeFileSync(entry, source);
  const config = buildConfig(consumerRoot, entry, path.resolve(outputDir), false);
  config.target = 'node';
  const stats = await new Promise((resolve, reject) => {
    const compiler = webpack(config);
    compiler.run((error, result) => compiler.close(closeError => error || closeError ? reject(error || closeError) : resolve(result)));
  });
  const detail = stats.toJson(statsOptions());
  fs.mkdirSync(outputDir,{recursive:true});
  fs.writeFileSync(path.join(outputDir, 'stats.json'), JSON.stringify(detail));
  if (stats.hasErrors() || detail.warnings.some(w => /export .*was not found|Module not found/.test(w.message))) throw new Error(`Production runtime bundle ${scenario}/${variant}: ${[...detail.errors, ...detail.warnings].map(w => w.message).join('\n')}`);
  assertToolkitOrigin(detail, consumerRoot);
  const bundlePath = path.join(outputDir, 'entry.js');
  return { scenario, variant, bundlePath, fixtureSha256: hash(source), bundleSha256: hash(fs.readFileSync(bundlePath)), emittedModules: collectEmittedModules(detail, consumerRoot), warnings: detail.warnings.map(w => w.message) };
}

function assertToolkitOrigin(detail, consumerRoot) {
  assertToolkitModuleOrigins(detail,path.join(consumerRoot,'node_modules',toolkit));
  collectEmittedModules(detail,consumerRoot); // Reject incomplete concatenated graphs.
}
function executeProbe(bundlePath, scenario, consumerRoot) {
  // The child resolves all React externals through the installed consumer.
  const result = spawnSync(process.execPath, [__filename, '--run-probe', scenario, bundlePath, consumerRoot], { encoding: 'utf8', timeout: 30000, maxBuffer: 4 * 1024 * 1024, stdio: ['ignore', 'pipe', 'pipe'] });
  const metadata = { exit: result.status, signal: result.signal, stderr: result.stderr || '', failureMarkers: (result.stderr || '').split('\n').filter(line => /Error|Assertion|PRODUCTION_RUNTIME_FAILURE/.test(line)) };
  if(result.error || result.status !== 0 || metadata.failureMarkers.length) throw Object.assign(new Error(`Production probe ${scenario}: ${result.error?.message || metadata.stderr}`),metadata);
  return {...JSON.parse(result.stdout),...metadata};
}
function installHostAdapter(consumerRoot) {
  const consumerRequire=createRequire(path.join(consumerRoot,'package.json'));
  const original=Module._load;
  global.MessageChannel=window.MessageChannel;
  // Resolve first to avoid recursion in the external adapter.
  const host={};
  for(const name of ['react','react-dom','react-dom/test-utils']) host[name]=original.call(Module,consumerRequire.resolve(name),module,false);
  Module._load=function(request,parent,isMain) {
    if(host[request]) return host[request];
    if (/^(react|react-dom)(\/|$)/.test(request)) return original.call(this,consumerRequire.resolve(request),parent,isMain);
    if(request === '@microsoft/sp-component-base') return {ThemeProvider:{serviceKey:'production-theme'}};
    return original.call(this,request,parent,isMain);
  };
  return {React:host.react,ReactDOM:host['react-dom'],act:callback=>host['react-dom/test-utils'].act(()=>{callback();})};
}
function jsonResponse(value) { return new Response(JSON.stringify(value), { headers: { 'Content-Type': 'application/json' } }); }
function batchResponse(scenario, partial = false) {
  const result = scenario === 'list' ? { Id: 11 } : { Title: 'Standalone web' };
  const extra = partial ? ['--batchresponse_production', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', 'HTTP/1.1 403 Forbidden', 'Content-Type: application/json', '', JSON.stringify({error:{message:'Rejected item'}})] : [];
  return new Response(['--batchresponse_production', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', 'HTTP/1.1 200 OK', 'Content-Type: application/json', '', JSON.stringify(result), ...extra, '--batchresponse_production--', ''].join('\r\n'), { headers: { 'Content-Type': 'multipart/mixed; boundary=batchresponse_production' } });
}
async function runProbe(scenario, bundlePath, consumerRoot) {
  if (!scenarios.includes(scenario)) throw new Error('Unknown production runtime scenario');
  const {JSDOM}=require('jsdom');
  const dom=new JSDOM('<!doctype html><html><head></head><body></body></html>',{url:base});
  global.window=dom.window;global.document=dom.window.document;
  Object.defineProperty(global,'navigator',{value:dom.window.navigator,configurable:true});
  global.HTMLElement=dom.window.HTMLElement;global.localStorage=dom.window.localStorage;global.sessionStorage=dom.window.sessionStorage;
  window.requestAnimationFrame=callback=>setTimeout(callback,0);window.cancelAnimationFrame=clearTimeout;
  if(scenario==='providers'||scenario==='styles') {
    const host=installHostAdapter(consumerRoot);
    const observation=await require(bundlePath).run({...host,assert});
    dom.window.close();
    return {scenario,pid:process.pid,assertions:observation.assertions,observation};
  }
  const calls = [];
  const result = await require(bundlePath).run(async (url, init) => {
    calls.push({ url, method: init.method, headers: init.headers, body: init.body });
    if (url === base + '/_api/$batch') return batchResponse(scenario, calls.filter(c=>c.url === base + '/_api/$batch').length>1);
    if (url === base + '/_api/contextinfo') return jsonResponse({ FormDigestValue: 'production-digest', FormDigestTimeoutSeconds: 1800 });
    if (scenario === 'context') return jsonResponse({ Title: 'Standalone web' });
    if(scenario==='list' && init.method==='POST' && !init.headers['X-HTTP-Method']) return jsonResponse({Id:12});
    if(scenario==='list' && /items\(7\)$/.test(url)) return jsonResponse({Id:7,Title:'Standalone task'});
    if (scenario === 'list') return jsonResponse({ value: [{ Id: 7, Title: 'Standalone task' }] });
    if (scenario === 'search' && url.includes('/suggest')) return jsonResponse({ Queries: ['Standalone suggestion'], PeopleNames: [], PersonalResults: [] });
    if (scenario === 'search') return jsonResponse({ PrimaryQueryResult: { RelevantResults: { RowCount: 1, TotalRows: 3, TotalRowsIncludingDuplicates: 3, Table: { Rows: [{ Cells: [{ Key: 'DocId', Value: '9', ValueType: 'Edm.Int64' }, { Key: 'Title', Value: 'Standalone search', ValueType: 'Edm.String' }, { Key: 'Rank', Value: '12', ValueType: 'Edm.Double' }] }] } }, RefinementResults: { Refiners: [] } } });
    throw new Error('Unexpected HTTP request: ' + url);
  });
  if (scenario === 'context') {
    assert.deepEqual(result, { web: { Title: 'Standalone web' }, targeted:{Title:'Standalone web'}, resolved: 'https://tenant.test/sites/other' });
    assert.deepEqual(calls.map(c => c.url), [base + '/_api/web','https://tenant.test/sites/other/_api/web']);
    assert.equal(calls[0].headers['X-Standalone'], 'context');
  } else if (scenario === 'list') {
    assert.deepEqual(result.query, { items: [{ Id: 7, Title: 'Standalone task' }], effectivePageSize: 2, hasMore: false, nextSkip: 1 });
    assert.deepEqual(result.batch, { value: [11], errors: [], summaryError: undefined });
    assert.deepEqual(calls.slice(0,2).map(c => c.url), [base + "/_api/web/lists/getByTitle('Tasks')/items?%24top=2", base + '/_api/$batch']);
    assert.equal(result.created,12);assert.equal(result.loaded.Id,7);
    assert.deepEqual(result.partial.value,[11]);assert.equal(result.partial.errors.length,1);assert.match(result.partial.errors[0],/403|Rejected/);assert.ok(result.partial.summary);
    const urls=calls.slice(2,6).map(c=>decodeURIComponent(c.url));
    assert.ok(urls[0].includes("getByTitle('Owner''s tasks')"));
    assert.ok(urls[1].includes("abcdefab-1234-1234-abcd-1234567890ab"),urls[1]);
    assert.ok(urls[2].includes("getList('/sites/standalone/Lists/Owner''s tasks')"));assert.equal(urls[2],urls[3]);
    assert.ok(calls.some(c=>c.headers['X-HTTP-Method']==='MERGE'));assert.ok(calls.some(c=>c.headers['X-HTTP-Method']==='DELETE'));
    assert.equal(calls[0].method, 'GET');
    assert.equal(calls[1].method, 'POST');
    assert.match(calls[1].body, /POST https:\/\/tenant\.test\/sites\/standalone\/_api\/web\/lists\/getByTitle\('Tasks'\)\/items HTTP\/1\.1/);
    assert.match(calls[1].body, /"Title":"Created task"/);
  } else if (scenario === 'search') {
    assert.equal(result.search.totalResults, 3);
    assert.deepEqual(result.search.refiners, []);
    assert.equal(result.search.results.length, 1);
    assert.equal(result.search.results[0].id, '9');
    assert.equal(result.search.results[0].data.Title, 'Standalone search');
    assert.equal(result.search.results[0].rank, 12);
    assert.deepEqual(result.suggestions, ['Standalone suggestion']);
    assert.deepEqual(calls.map(c => c.url), [base + '/_api/search/postquery', base + '/_api/search/suggest?querytext=%27Standalone%27']);
    assert.equal(calls[0].method, 'POST');
    assert.equal(JSON.parse(calls[0].body).request.Querytext, 'Standalone');
    assert.equal(JSON.parse(calls[0].body).request.RowLimit, 2);
  } else {
    assert.deepEqual(result, { Title: 'Standalone web' });
    assert.deepEqual(calls.map(c => c.url), [base + '/_api/$batch']);
    assert.equal(calls[0].method, 'POST');
    assert.match(calls[0].body, /GET https:\/\/tenant\.test\/sites\/standalone\/_api\/web HTTP\/1\.1/);
  }
  dom.window.close();
  return { scenario, calls, result, pid:process.pid, assertions:[`${scenario}: real bundled HTTP operations and parsed results`] };
}

async function verifyProductionRuntime({ consumerRoot, outputDir }) {
  consumerRoot = fs.realpathSync(consumerRoot); outputDir = path.resolve(outputDir);
  const toolkitDir = path.join(consumerRoot, 'node_modules', toolkit);
  if (fs.realpathSync(toolkitDir) !== toolkitDir || fs.lstatSync(toolkitDir).isSymbolicLink()) throw new Error('Production runtime requires an installed toolkit, not a workspace link');
  const manifest = JSON.parse(fs.readFileSync(path.join(consumerRoot, 'package.json')));
  const dependency = manifest.dependencies?.[toolkit];
  if (!dependency?.startsWith('file:') || !dependency.endsWith('.tgz')) throw new Error('Production runtime requires consumer file tarball provenance');
  const tarball = path.resolve(consumerRoot, dependency.slice(5));
  const report = { schemaVersion: 1, complete: false, passed: false, provenance: { ...readRuntimeToolchain(), consumerRoot, toolkitDir, tarballSha256: hash(fs.readFileSync(tarball)), toolkitManifestSha256: hash(fs.readFileSync(path.join(toolkitDir, 'package.json'))), interception: 'HTTP send only; client and registration bundled in one realm; fresh process per probe' }, probes: [], warnings: [], failures: [] };
  fs.mkdirSync(outputDir, { recursive: true });
  const save = () => fs.writeFileSync(path.join(outputDir, 'runtime-report.json'), JSON.stringify(report, null, 2) + '\n');
  save();
  const start=Date.now();
  await runRuntimeMatrix({report,cases:scenarios.flatMap(scenario=>variants.map(variant=>({scenario,variant}))),
    compile:probe=>compileProbe({consumerRoot,outputDir:path.join(outputDir,probe.scenario,probe.variant),...probe}),
    execute:(bundlePath,scenario)=>executeProbe(bundlePath,scenario,consumerRoot),save});
  report.durationMs=Date.now()-start;save();
  return report;
}
async function runRuntimeMatrix({report,cases,compile,execute,save}) {
  for(const probe of cases) {
    let bundle;
    try {
      bundle=await compile(probe);
      report.warnings.push(...bundle.warnings.map(message=>({...probe,message})));
      const observation=await execute(bundle.bundlePath,probe.scenario);
      report.probes.push({...bundle,...probe,status:'passed',exit:observation.exit,signal:observation.signal,assertions:observation.assertions,failureMarkers:observation.failureMarkers,observation});
    } catch(error) {
      const failure={...probe,message:error.message,exit:error.exit??null,signal:error.signal??null,stderr:error.stderr?.toString(),failureMarkers:error.failureMarkers??['BUILD_OR_EXECUTION_FAILURE']};
      report.probes.push({...bundle,...failure,status:'failed'});report.failures.push(failure);
    }
    save();
  }
  report.complete=true;report.passed=report.failures.length===0 && report.probes.length===cases.length;save();
}
if (require.main === module) {
  if (process.argv[2] === '--run-probe') runProbe(process.argv[3], path.resolve(process.argv[4]), path.resolve(process.argv[5])).then(result => console.log(JSON.stringify(result))).catch(error => { console.error('PRODUCTION_RUNTIME_FAILURE',error); process.exitCode = 1; });
  else {
    const args = process.argv.slice(2); const options = {};
    for (let i = 0; i < args.length; i += 2) {
      if (!['--consumer-root', '--output'].includes(args[i]) || !args[i + 1]) throw new Error('Required: --consumer-root PATH --output PATH');
      options[args[i]] = args[i + 1];
    }
    if (!options['--consumer-root'] || !options['--output']) throw new Error('Required: --consumer-root PATH --output PATH');
    verifyProductionRuntime({ consumerRoot: options['--consumer-root'], outputDir: options['--output'] }).then(report => { console.log(`production runtime: ${report.probes.length} probes; passed=${report.passed}; warnings=${report.warnings.length}`); if (!report.passed) process.exitCode = 1; }).catch(error => { console.error(error); process.exitCode = 1; });
  }
}
module.exports = { verifyProductionRuntime, compileProbe, executeProbe, runRuntimeMatrix, assertToolkitOrigin };
