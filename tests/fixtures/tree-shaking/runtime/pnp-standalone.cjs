// Source regression only: production-package scenarios use a separate bundled fixture.
// Each invocation owns its loader cache; no root barrel or shared PnP harness is loaded.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');

const root = path.resolve(__dirname, '../../../../packages/spfx-react-toolkit/src');
const cache = new Map();
function load(file) {
  if (cache.has(file)) return cache.get(file).exports;
  const module = { exports: {} };
  cache.set(file, module);
  const source = ts.transpileModule(fs.readFileSync(file, 'utf8'), {
    compilerOptions: { allowJs: true, module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 }
  }).outputText;
  function sourceRequire(specifier) {
    if (specifier.startsWith('@pnp/')) return load(require.resolve(specifier));
    if (!specifier.startsWith('.')) return require(specifier);
    const resolved = path.resolve(path.dirname(file), specifier);
    const dependency = ['', '.ts', '.js'].map(ext => resolved + ext)
      .find(candidate => fs.existsSync(candidate) && fs.statSync(candidate).isFile());
    if (!dependency) throw new Error('Unresolved source import: ' + specifier);
    return load(dependency);
  }
  vm.runInThisContext('(function(require, module, exports) {' + source + '\n})', { filename: file })(sourceRequire, module, module.exports);
  return module.exports;
}

const base = 'https://tenant.test/sites/standalone';
const calls = [];
function intercept(respond) {
  return instance => {
    instance.on.send.replace(async (url, init) => {
      const call = { url: String(url), init };
      calls.push(call);
      const result = respond(call);
      return result instanceof Response ? result : new Response(JSON.stringify(result), {
        headers: { 'Content-Type': 'application/json' }
      });
    });
    return instance;
  };
}
function bareClient(respond) {
  const { spfi } = load(require.resolve('@pnp/sp/fi'));
  const { DefaultInit, DefaultHeaders } = load(require.resolve('@pnp/sp/behaviors/defaults'));
  const { DefaultParse } = load(require.resolve('@pnp/queryable'));
  return spfi(base).using(DefaultInit(), DefaultHeaders(), DefaultParse(), intercept(respond));
}
function service(name) {
  return load(path.join(root, 'services', name + '.ts'));
}

function batchResponse() {
  return new Response([
    '--batchresponse_standalone', 'Content-Type: application/http',
    'Content-Transfer-Encoding: binary', '', 'HTTP/1.1 200 OK',
    'Content-Type: application/json', '', '{"Title":"Standalone web"}',
    '--batchresponse_standalone--', ''
  ].join('\r\n'), { headers: { 'Content-Type': 'multipart/mixed; boundary=batchresponse_standalone' } });
}

async function main(scenario) {
  if (scenario === 'context') {
    // No explicit webs or batching registration: both belong to the context leaf.
    const { createSPFxPnPContextService } = service('spfx-pnp-context.service');
    const pageContext = { web: { absoluteUrl: base } };
    const context = createSPFxPnPContextService({ pageContext }, pageContext);
    const sp = context.createSPFI(undefined, { headers: { 'X-Standalone': 'context' } })
      .using(intercept(call => {
        if (call.url.endsWith('/$batch')) return batchResponse();
        if (call.url.endsWith('/contextinfo')) return { FormDigestValue: 'standalone-digest', FormDigestTimeoutSeconds: 1800 };
        return { Title: 'Standalone web' };
      }));
    assert.equal(context.resolveSiteUrl('/sites/other/'), 'https://tenant.test/sites/other');
    assert.deepEqual(await sp.web(), { Title: 'Standalone web' });
    const [batched, execute] = sp.batched();
    const pending = batched.web();
    await execute();
    assert.deepEqual(await pending, { Title: 'Standalone web' });
    assert.deepEqual(calls.map(call => call.url), [base + '/_api/web', base + '/_api/contextinfo', base + '/_api/$batch']);
    assert.equal(calls[0].init.headers['X-Standalone'], 'context');
    assert.equal(calls[2].init.headers['X-RequestDigest'], 'standalone-digest');
    assert.match(calls[2].init.body, /GET https:\/\/tenant\.test\/sites\/standalone\/_api\/web HTTP\/1\.1/);
  } else if (scenario === 'batch') {
    const sp = bareClient(() => batchResponse());
    assert.equal(typeof sp.batched, 'undefined', 'Caller setup must not register toolkit batching');
    load(require.resolve('@pnp/sp/webs'));
    assert.equal(typeof sp.batched, 'undefined', 'Caller webs must not register batching');
    const { createSPFxPnPService } = service('spfx-pnp.service');
    // The caller owns webs used in its callback, while the leaf owns batching.
    const result = await createSPFxPnPService(sp).batch(batched => batched.web());
    assert.deepEqual(result, { Title: 'Standalone web' });
    assert.equal(calls.length, 1);
    assert.equal(calls[0].url, base + '/_api/$batch');
    assert.equal(calls[0].init.method, 'POST');
    assert.match(calls[0].init.body, /GET https:\/\/tenant\.test\/sites\/standalone\/_api\/web HTTP\/1\.1/);
  } else if (scenario === 'search') {
    const sp = bareClient(call => call.url.includes('/suggest') ? {
      Queries: ['Standalone suggestion'], PeopleNames: [], PersonalResults: []
    } : {
      PrimaryQueryResult: {
        RelevantResults: {
          RowCount: 1, TotalRows: 3, TotalRowsIncludingDuplicates: 3,
          Table: { Rows: [{ Cells: [
            { Key: 'DocId', Value: '9', ValueType: 'Edm.Int64' },
            { Key: 'Title', Value: 'Standalone search', ValueType: 'Edm.String' },
            { Key: 'Rank', Value: '12', ValueType: 'Edm.Double' }
          ] }] }
        },
        RefinementResults: { Refiners: [] }
      }
    });
    const { createSPFxPnPSearchService } = service('spfx-pnp-search.service');
    const search = createSPFxPnPSearchService(sp, { pageSize: 2 });
    const result = await search.search('Standalone');
    assert.equal(result.totalResults, 3);
    assert.deepEqual(result.refiners, []);
    assert.equal(result.results.length, 1);
    assert.equal(result.results[0].id, '9');
    assert.equal(result.results[0].data.Title, 'Standalone search');
    assert.equal(result.results[0].rank, 12);
    assert.deepEqual(await search.suggest('Standalone'), ['Standalone suggestion']);
    assert.equal(calls.length, 2);
    assert.equal(calls[0].url, base + '/_api/search/postquery');
    assert.equal(calls[0].init.method, 'POST');
    const request = JSON.parse(calls[0].init.body).request;
    assert.equal(request.Querytext, 'Standalone');
    assert.equal(request.RowLimit, 2);
    assert.equal(calls[1].url, base + '/_api/search/suggest?querytext=%27Standalone%27');
  } else {
    if (scenario !== 'list') throw new Error('Unknown scenario: ' + scenario);
    const sp = bareClient(() => ({ value: [{ Id: 7, Title: 'Standalone task' }] }));
    const { createSPFxPnPListService } = service('spfx-pnp-list.service');
    const result = await createSPFxPnPListService(sp, 'Tasks', 2).query();
    assert.deepEqual(result, {
      items: [{ Id: 7, Title: 'Standalone task' }], effectivePageSize: 2, hasMore: false, nextSkip: 1
    });
    assert.equal(calls.length, 1);
    assert.equal(calls[0].url, base + "/_api/web/lists/getByTitle('Tasks')/items?%24top=2");
    assert.equal(calls[0].init.method, 'GET');
  }
  console.log(JSON.stringify({ scenario, urls: calls.map(call => call.url) }));
}
main(process.argv[2]).catch(error => { console.error(error); process.exitCode = 1; });
