const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const ts = require('typescript');
const { loadHook } = require('./async-test-harness.cjs');
const { createHarness } = require('./services-test-harness.cjs');
const { buildApiGraph } = require('../scripts/documentation-model.cjs');
const { buildDocumentationModel, getDocumentationExample } = require('../scripts/documentation-contract.cjs');
const config = require('../scripts/documentation.config.cjs');
let model;
test.before(() => { model = buildDocumentationModel(buildApiGraph(path.resolve(__dirname, '..'), config), config); });

function compile(example, publicApi, React) {
  const source = ts.createSourceFile('example.tsx', example.code, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  const component = source.statements.find(node => ts.isFunctionDeclaration(node));
  assert.ok(component?.name, 'normalized interaction recipe must declare its component');
  const output = ts.transpileModule(example.code + `\nexport { ${component.name.text} as Probe };`, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React },
  }).outputText;
  const module = { exports: {} };
  const boundary = specifier => {
    if (specifier === 'react') return React;
    if (specifier === '@apvee/spfx-react-toolkit') return publicApi;
    if (specifier === '@microsoft/sp-http') return Object.fromEntries(['HttpClient', 'AadHttpClient', 'SPHttpClient'].map(name => [name, { configurations: { v1: {} } }]));
    if (specifier.startsWith('@pnp/')) return {};
    throw new Error('Unexpected runtime example import: ' + specifier);
  };
  new Function('require', 'module', 'exports', output)(boundary, module, module.exports);
  return module.exports.Probe;
}
async function flush(h, action = () => {}) {
  await h.act(async () => { action(); for (let i = 0; i < 20; i++) await Promise.resolve(); });
  // Leave a real event-loop turn for Node's unhandledRejection detection.
  await new Promise(resolve => setImmediate(resolve));
}

test('normalized Graph interaction recipes catch actual invoke failures, gate readiness and render results', async t => {
  const examples = model.examples.filter(example => example.source.file.endsWith('/useSPFxMSGraphClient.ts') && example.code.includes('const run = async'));
  assert.equal(examples.length, 2, 'both invoke and readiness recipes are exercised');
  for (const example of examples) {
    const h = createHarness();
    t.after(() => h.close());
    const useAsyncInvoke = loadHook('useAsyncInvoke.internal', 'useAsyncInvoke');
    let ready = false, mode = 'reject', calls = 0;
    const request = async () => { calls++; if (mode === 'reject') throw new Error('Expected Graph failure'); return { displayName: 'Ada' }; };
    const client = { api: () => ({ select: () => ({ get: request }), get: request }) };
    const publicApi = { useSPFxMSGraphClient: () => ({ ...useAsyncInvoke(ready ? client : undefined), isReady: ready, isInitializing: !ready, initError: undefined }) };
    const Component = compile(example, publicApi, h.React);
    let mounted = h.mountComponent(Component);
    assert.equal(mounted.container.querySelector('button')?.disabled, true, 'unready client must disable the request');
    await flush(h, () => mounted.container.querySelector('button').click());
    assert.equal(calls, 0);
    mounted.unmount();
    ready = true;
    mounted = h.mountComponent(Component);
    await flush(h, () => mounted.container.querySelector('button').click());
    assert.match(mounted.container.textContent, /Expected Graph failure/);
    mode = 'success';
    await flush(h, () => mounted.container.querySelector('button').click());
    assert.match(mounted.container.textContent, /Ada/);
    assert.doesNotMatch(mounted.container.textContent, /Expected Graph failure/);
    mounted.unmount();
    h.close();
  }
});

// These are runtime boundaries for the edited recipes, not compiler declarations
// or native/tenant implementations. The Graph test above uses the real invoke hook.
test('the bounded corrected interaction recipes catch rejecting actions and render successful returned data', async t => {
  const files = new Set(['useSPFxPnPListByPath.ts', 'useSPFxSiteKeyValueStore.ts', 'useSPFxPnPListByUrl.ts', 'useSPFxTenantKeyValueStore.ts', 'useSPFxPerformance.ts', 'useSPFxAadTokenProvider.ts', 'useSPFxHttpClient.ts', 'useSPFxPnP.ts', 'useSPFxAadHttpClient.ts', 'useSPFxPnPListById.ts', 'useSPFxPnPSearch.ts', 'useSPFxSPHttpClient.ts', 'useSPFxMSGraphClient.ts', 'useSPFxPnPList.ts']);
  const examples = model.examples.filter(example => files.has(path.posix.basename(example.source.file)) && example.code.includes('setFailure(failure instanceof Error'));
  assert.equal(examples.length, 30, 'all physical corrected recipes must be executed');
  for (const example of examples) {
    const h = createHarness();
    t.after(() => h.close());
    let reject = true, calls = 0;
    const operation = async value => { calls++; await Promise.resolve(); if (reject) throw new Error('Expected boundary failure'); return value; };
    function useBoundary() {
      const [items, setItems] = h.React.useState([]);
      const query = async () => { const next = await operation([{ Id: 1, Title: 'Ada' }]); setItems(next); return next; };
      return {
        isReady: true, isLoading: false, isWriting: false, isInitialized: true,
        isInitializing: false, loading: false, error: undefined, initError: undefined,
        items, refiners: [{ name: 'FileType', entries: [{ token: 'word', value: 'Word', count: 1 }] }],
        resourceUrl: 'https://api.example.test', baseUrl: 'https://tenant.example.test',
        invoke: () => operation({ displayName: 'Ada' }), batch: () => operation([{ Title: 'Ada' }]),
        query, search: () => operation([{ id: '1', data: { Title: 'Ada' } }]),
        suggest: () => operation(['Ada']), applyRefiner: () => operation(undefined),
        get: () => operation({ value: true }), save: () => operation(undefined),
        time: () => operation({ durationMs: 12, result: { status: 200 } }),
        tokenProvider: { getToken: () => operation('boundary-token-not-rendered') },
      };
    }
    const hooks = [...example.code.matchAll(/\b(useSPFx\w+)\b/g)].map(match => match[1]);
    const publicApi = Object.fromEntries(hooks.map(hook => [hook, useBoundary]));
    publicApi.decodeSPFxJwtPayload = () => ({ scp: 'User.Read' });
    const mounted = h.mountComponent(compile(example, publicApi, h.React));
    const label = `${example.source.file}:${example.source.line}`;
    await flush(h, () => mounted.container.querySelector('button').click());
    assert.ok(calls > 0, label + ': rejecting action reached');
    assert.match(mounted.container.textContent, /Expected boundary failure/, label);
    reject = false;
    await flush(h, () => mounted.container.querySelector('button').click());
    assert.doesNotMatch(mounted.container.textContent, /Expected boundary failure/, label);
    assert.match(mounted.container.querySelector('pre').textContent, /Ada|true|Saved settings|Filter applied|200|User.Read/, label + ': returned result visible');
    assert.doesNotMatch(mounted.container.textContent, /boundary-token-not-rendered/, label + ': raw access tokens stay out of UI');
    mounted.unmount();
    h.close();
  }
});

// Select the residual public recipes by their canonical owner/title, independent
// of whether their handlers contain a catch or any particular implementation.
for (const [reference, title, kind] of [
  ['SearchVerticals', 'Search People with an explicit builder', 'people'],
  ['useSPFxPnPSearch', 'Basic usage', 'documents'],
]) test(`normalized ${reference} ${title ?? 'example'} handles search failure and renders its result`, async t => {
  const example = getDocumentationExample(model, reference, title);
  const h = createHarness();
  t.after(() => h.close());
  let reject = true, loading = false, requests = 0;
  const peopleSource = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';
  const search = async query => {
    requests++;
    if (kind === 'people') {
      const builder = {
        text(value) { assert.equal(value, 'john'); return this; },
        sourceId(value) { assert.equal(value, peopleSource); return this; },
      };
      assert.equal(query(builder), builder, 'the People builder remains a composed query');
    } else assert.equal(query, 'ContentType:Document');
    await Promise.resolve();
    if (reject) throw new Error('Expected search failure');
    return [{ id: 'person-or-document', data: { Title: 'Ada' } }];
  };
  const publicApi = {
    SearchVerticals: { People: peopleSource },
    useSPFxPnPSearch: () => ({ search, loading, error: undefined, results: [] }),
  };
  const Component = compile(example, publicApi, h.React);
  let mounted = h.mountComponent(Component);
  await flush(h, () => mounted.container.querySelector('button').click());
  assert.match(mounted.container.textContent, /Expected search failure/);
  reject = false;
  await flush(h, () => mounted.container.querySelector('button').click());
  assert.doesNotMatch(mounted.container.textContent, /Expected search failure/);
  assert.match(mounted.container.querySelector('pre').textContent, /Ada/);
  mounted.unmount();
  loading = true;
  mounted = h.mountComponent(Component);
  assert.equal(mounted.container.querySelector('button').disabled, true);
  const before = requests;
  await flush(h, () => mounted.container.querySelector('button').click());
  assert.equal(requests, before, 'pending search prevents another event request');
});
