const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const hostFile = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/NativePropertyPaneDemoHost.ts');
function hostClass() {
  assert.ok(fs.existsSync(hostFile), 'private native protocol implementation exists');
  const module = { exports: {} };
  new Function('require', 'module', 'exports', ts.transpileModule(fs.readFileSync(hostFile, 'utf8'), {fileName:hostFile,compilerOptions:{target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.CommonJS}}).outputText)(name => name === '@microsoft/sp-property-pane' ? { PropertyPaneFieldType: { Custom: 1 } } : require(name), module, module.exports);
  return module.exports.NativePropertyPaneDemoHost;
}
test('native prepare is pure and own unproduced lease cancellation never cancels a successor', () => {
  const Host = hostClass();
  let open = false, factories = 0, refreshes = 0;
  const host = new Host({ instance: {}, readProperties: () => ({}), isOpen: () => open, refresh: () => { refreshes++; } });
  const create = () => { factories++; return { getConfiguration: () => ({ pages: [] }), fieldChanged() {}, configurationComplete() {}, changesApplied() {}, dispose() {} }; };
  const a = host.port.prepare(create); assert.equal(a.status, 'prepared'); a.lease.release(); a.lease.release();
  const b = host.port.prepare(create); a.lease.release(); assert.equal(b.status, 'prepared');
  assert.equal(factories, 0); assert.equal(refreshes, 0);
  host.getConfiguration(() => ({ pages: [] })); assert.equal(factories, 0);
  open = true; host.getConfiguration(() => ({ pages: [] })); assert.equal(factories, 1);
});

const { createPinnedSdk } = require('./fixtures/native-property-pane/sdk-schedule.cjs');
function fixture(options = {}) {
  const Host = hostClass(); const events = [], controllers = [], elements = new Set();
  let open = false, bag = { values: ['first'] }, alive = 0, peak = 0, refreshes = 0;
  const instance = {};
  const host = new Host({ instance, readProperties: () => bag, isOpen: () => options.sdk?.open ?? open, refresh: () => { refreshes++; } });
  function prepare(name, { failFactory = false, failConfiguration = false, fields = 1 } = {}) {
    return host.port.prepare(environment => {
      events.push('create:' + name); if (failFactory) throw new Error('Source factory');
      alive++; peak = Math.max(peak, alive);
      const record = { name, environment, fields: [], disposed: 0, changes: [], completed: 0, applied: 0, reads: [] }; controllers.push(record);
      return {
        getConfiguration: () => {
          record.reads.push(environment.readProperties()); if (failConfiguration) throw new Error('Source configuration');
          return { pages: [{ groups: [{ isGroupNameHidden: true, groupFields: Array.from({ length: fields }, (_, i) => {
            const field = { type: 1, targetProperty: 'values', properties: { key: 'subject-' + i,
              onRender: element => { elements.add(element); events.push('mount:' + name); },
              onDispose: element => { assert.equal(options.sdk?.open ?? open, true, 'latest native disposal precedes public close flag'); elements.delete(element); events.push('unmount:' + name); }
            } }; record.fields.push(field); return field;
          }) }] }] };
        },
        fieldChanged: (...args) => { record.changes.push(args); environment.refresh(); },
        configurationComplete: () => { record.completed++; environment.refresh(); },
        changesApplied: () => { record.applied++; environment.refresh(); },
        dispose: () => { assert.equal(record.disposed++, 0, 'one effective controller disposal'); alive--; events.push('dispose:' + name); }
      };
    });
  }
  const configuration = () => host.getConfiguration(() => ({ pages: [] }));
  return { host, prepare, configuration, events, controllers, elements, get alive() { return alive; }, get peak() { return peak; },
    get refreshes() { return refreshes; }, get open() { return open; }, set open(value) { open = value; }, set bag(value) { bag = value; },
    field(config = configuration(), index = 0) { return config.pages[0].groups[0].groupFields[index]; },
    close(...fields) { for (const [field, element] of fields) field.properties.onDispose(element); open = false; host.configurationComplete(); }
  };
}
test('live bag reads, first native witness, latest same-DOM disposal, scoped refresh and terminal invalidation', () => {
  const f = fixture(); const lease = f.prepare('A'); f.open = true;
  const first = f.field(), element = {};
  assert.equal(f.controllers[0].reads[0].values[0], 'first');
  f.host.fieldChanged('values', [], ['before witness']); assert.equal(f.controllers[0].changes.length, 0);
  first.properties.onRender(element); f.host.rendered();
  f.host.fieldChanged('values', [], ['accepted']); assert.equal(f.refreshes, 1);
  f.host.configurationComplete(); f.host.changesApplied(); f.controllers[0].environment.refresh();
  assert.equal(f.refreshes, 1, 'Start/Complete/AfterApply/background cannot refresh'); assert.equal(f.alive, 1);
  f.bag = { values: ['SDK replaced'] }; const second = f.field();
  assert.equal(f.controllers[0].reads.at(-1).values[0], 'SDK replaced');
  second.properties.onRender(element); f.host.rendered();
  first.properties.onDispose(element); assert.equal(f.elements.size, 1); assert.equal(f.events.filter(e => e === 'unmount:A').length, 0);
  lease.lease.release(); assert.equal(f.alive, 1, 'produced lease survives source navigation/reset');
  const cell = [...f.host.produced.fields][0];
  f.host.dispose(); f.host.dispose(); second.properties.onRender(element); second.properties.onDispose(element);
  assert.equal(f.elements.size, 0); assert.equal(f.alive, 0); assert.equal(f.controllers[0].disposed, 1); assert.equal(f.events.filter(e => e === 'unmount:A').length, 1);
  for (const key of ['controller', 'environment']) assert.equal(cell.generation?.[key], undefined);
  for (const key of ['render', 'dispose', 'dispatch', 'generation']) assert.equal(cell[key], undefined, 'retired small field cell clears ' + key);
  assert.equal(f.host.port.prepare(() => { throw new Error(); }).status, 'blocked');
});
test('open and remaining native DOM block prepare; Complete/foreign switch and bare parent cannot retire or grant ownership', () => {
  const f = fixture(); f.prepare('A'); f.open = true; const field = f.field(), element = {};
  field.properties.onRender(element); f.host.rendered();
  assert.equal(f.prepare('B').status, 'blocked');
  f.host.configurationComplete(); f.host.rendered();
  assert.equal(f.alive, 1); f.host.fieldChanged('values', 1, 2); assert.equal(f.controllers[0].changes.length, 1);
  f.open = false; assert.equal(f.prepare('B').status, 'blocked', 'resources remain even with global closed flag');
  f.open = true; f.close([field, element]); assert.equal(f.alive, 0); assert.equal(f.prepare('B').status, 'prepared');
  const stale = field; f.open = true; f.field(); stale.properties.onRender(element); f.host.rendered();
  f.host.fieldChanged('values', 2, 3); assert.equal(f.controllers[1].changes.length, 0, 'old witnesses cannot grant new generation rights'); f.host.dispose();
});
test('source factory/configuration failure is transactional and retains the previous cold owner without partial publication', () => {
  for (const failure of ['failFactory', 'failConfiguration']) {
    const f = fixture(); f.prepare('A'); f.open = true; const old = f.configuration();
    f.open = false; f.host.configurationComplete(); f.prepare('bad', { [failure]: true }); f.open = true;
    const restored = f.configuration(); assert.equal(restored.pages.length, 1); assert.equal(f.host.produced.controller !== undefined, true);
    assert.equal(f.controllers[0].disposed, 0); assert.equal(f.alive, 1); assert.equal(f.peak, failure === 'failFactory' ? 1 : 2);
    assert.equal(f.host.candidate, undefined); assert.equal(restored, old); f.host.dispose(); assert.equal(f.alive, 0);
  }
});
test('pinned cold B/C/D uses one shared resource and mutable latest DTO with strict constructor peak two', async () => {
  const sdk = createPinnedSdk(), f = fixture({ sdk });
  f.prepare('B'); const b = sdk.edit('one', f.host); sdk.resource('one').resolve(); await sdk.drain(); assert.equal(f.alive, 1); assert.deepEqual(sdk.trace, ['G:one', 'I:one']);
  sdk.close(f.host); f.prepare('C'); const c = sdk.edit('one', f.host); await sdk.drain(); assert.equal(f.alive, 1); assert.equal(f.controllers[0].disposed, 1);
  sdk.close(f.host); f.prepare('D'); const d = sdk.edit('one', f.host); await sdk.drain(); assert.equal(f.alive, 1); assert.equal(f.controllers[1].disposed, 1); assert.equal(f.peak, 2);
  sdk.load(); await Promise.all([b, c, d]); assert.equal(f.events.filter(e => e.startsWith('mount:')).every(e => e === 'mount:D'), true);
  assert.equal(f.controllers[0].changes.length, 0); f.host.fieldChanged('values', 0, 1); assert.equal(f.controllers[2].changes.length, 1);
  sdk.close(f.host); assert.equal(f.alive, 0); assert.equal(f.controllers.every(c => c.disposed === 1), true);
});
test('pinned cold chunk before next Edit witnesses withdrawn subject without mounting and routes stale built-ins until successor commit', async () => {
  const sdk = createPinnedSdk(), f = fixture({ sdk }); f.prepare('B'); const b = sdk.edit('one', f.host); sdk.resource('one').resolve(); await sdk.drain(); sdk.close(f.host);
  const oldFields = [...f.host.produced.fields]; f.prepare('C'); sdk.load(); await b;
  assert.equal(f.events.filter(e => e.startsWith('mount:')).length, 0); assert.equal(f.host.nativeView.controller !== undefined, true);
  const c = sdk.edit('one', f.host); f.host.fieldChanged('values', 'B', 'while B visible');
  assert.equal(f.controllers[0].changes.length, 1);
  await c; assert.equal(f.controllers[0].disposed, 1); assert.equal(f.events.filter(e => e === 'mount:C').length, 1); assert.equal(f.peak, 2);
  for (const cell of oldFields) for (const key of ['render', 'dispose', 'dispatch', 'generation']) assert.equal(cell[key], undefined);
  sdk.close(f.host); assert.equal(f.alive, 0);
});
test('cached pending resources execute each getter against the current candidate and foreign ID gates installation', async () => {
  const sdk = createPinnedSdk(), a = fixture({ sdk }), b = fixture({ sdk });
  a.prepare('old'); const p1 = sdk.edit('A', a.host); sdk.close(a.host); a.prepare('current'); const p2 = sdk.edit('A', a.host);
  sdk.close(a.host); b.prepare('foreign'); const p3 = sdk.edit('B', b.host); sdk.resource('A').resolve(); await sdk.drain();
  assert.equal(a.controllers.length, 1); assert.equal(a.controllers[0].name, 'current'); assert.equal(sdk.trace.filter(e => e === 'foreign:A').length, 2);
  sdk.resource('B').resolve(); await sdk.drain(); sdk.load(); await Promise.all([p1, p2, p3]);
  assert.equal(sdk.data.id, 'B'); assert.equal(b.controllers.length, 1); assert.equal(a.elements.size, 0); assert.equal(b.elements.size, 1);
  sdk.close(b.host); a.host.dispose(); assert.equal(a.alive, 0); assert.equal(b.alive, 0);
});
test('warm pinned component commits in the microtask drain before the next real Close task', async () => {
  const sdk = createPinnedSdk(), f = fixture({ sdk }); sdk.load(); sdk.resource('one').resolve();
  f.prepare('A'); await sdk.edit('one', f.host); assert.equal(sdk.warm, true); sdk.close(f.host); assert.equal(f.controllers[0].disposed, 1);
  f.prepare('B'); const next = sdk.edit('one', f.host); await next; assert.equal(f.controllers[1].disposed, 0);
  sdk.close(f.host); assert.equal(f.alive, 0); assert.ok(f.events.indexOf('mount:B') < f.events.indexOf('unmount:B')); assert.equal(f.peak, 1);
});
test('cold successful and external failed DTO paths have identical public observations; automatic recovery is unsupported', async () => {
  async function run(failDto) {
    const sdk = createPinnedSdk(), f = fixture({ sdk }); f.prepare('B'); const b = sdk.edit('one', f.host); sdk.resource('one').resolve(); await sdk.drain(); sdk.close(f.host);
    f.prepare('C'); const c = sdk.edit('one', f.host, { failDto }); const caught = c.catch(error => error.message); await sdk.drain(); sdk.close(f.host);
    const observations = { creates: f.controllers.map(c => c.name), disposals: f.controllers.map(c => c.disposed), open: sdk.open, mounts: f.elements.size, live: f.alive };
    assert.equal(sdk.data.configuration.pages[0].groups.length, 1); const installed = sdk.data.configuration;
    const dataCell = installed.pages[0].groups[0].groupFields[0];
    f.host.dispose(); sdk.load(); await Promise.all([b, caught]); return { observations, activeDescriptor: dataCell.properties.onRender, trace: sdk.trace };
  }
  const success = await run(false), failure = await run(true);
  assert.deepEqual(success.observations, failure.observations);
  assert.ok(failure.trace.includes('DTO-throw:one')); assert.equal(success.trace.includes('DTO-throw:one'), false);
  // This pair exposes the undetectable failed-install boundary. No recovery PASS,
  // fake acknowledgement, timing settlement, or SDK failure swallowing is claimed.
});

test('mixed or absent witnesses cannot switch native ownership; later named collapse retains the same committed business owner', () => {
  const sdk = createPinnedSdk(), f = fixture({ sdk });
  // Build the valid cold/hidden path, then use controlled field witnesses to
  // isolate ambiguity handling independently of the Promise schedule tests.
  f.prepare('A');
  return (async () => {
    const pending = sdk.edit('one', f.host); sdk.resource('one').resolve(); await sdk.drain(); sdk.close(f.host); f.prepare('B'); sdk.load(); await pending;
    const old = f.host.nativeView, oldField = sdk.data.configuration.pages[0].groups[0].groupFields[0], next = sdk.edit('one', f.host);
    const producedField = f.field(), manualDom = {};
    oldField.properties.onRender({}); producedField.properties.onRender(manualDom);
    f.host.rendered(); assert.equal(f.host.nativeView, old, 'real mixed wrappers cannot acknowledge a new view');
    producedField.properties.onDispose(manualDom);
    // A bare/repeated parent supplies no new owner.
    f.host.rendered(); assert.equal(f.host.nativeView, old);
    await next; const current = f.host.nativeView; assert.notEqual(current, old);
    f.host.rendered(); f.host.rendered(); assert.equal(f.host.nativeView, current);
    sdk.close(f.host); assert.equal(f.host.nativeView, undefined); assert.equal(f.alive, 0);
  })();
});
test('two adapters keep own properties, callbacks, tokens and synchronous refresh scopes independent', () => {
  const a = fixture(), b = fixture(); a.prepare('A'); b.prepare('B'); a.open = true; b.open = true;
  const af = a.field(), bf = b.field(), ad = {}, bd = {};
  af.properties.onRender(ad); a.host.rendered(); bf.properties.onRender(bd); b.host.rendered();
  a.host.fieldChanged('values', 1, 2); assert.equal(a.refreshes, 1); assert.equal(b.refreshes, 0); assert.equal(b.controllers[0].changes.length, 0);
  a.host.dispose(); af.properties.onRender(ad); assert.equal(a.alive, 0); assert.equal(b.alive, 1);
  b.host.fieldChanged('values', 2, 3); assert.equal(b.refreshes, 1); b.host.dispose(); assert.equal(b.alive, 0);
});

test('a genuine named-group collapse releases its native mount but can remount the same descriptor while business ownership stays committed', () => {
  const f = fixture(); f.prepare('A'); f.open = true; const field = f.field(), first = {}, second = {};
  field.properties.onRender(first); f.host.rendered(); const business = f.host.nativeView;
  field.properties.onDispose(first); field.properties.onDispose(first); f.host.rendered();
  assert.equal(f.elements.size, 0); assert.equal(f.host.nativeView, business); assert.equal(f.alive, 1);
  field.properties.onRender(second); f.host.rendered(); assert.equal(f.elements.size, 1); assert.equal(f.host.nativeView, business);
  f.close([field, second]); assert.equal(f.alive, 0); assert.equal(f.events.filter(e => e === 'unmount:A').length, 2);
});
test('a source configuration throw during same-generation refresh preserves the last complete native descriptor', () => {
  const Host = hostClass(); let open = false, fail = false, disposed = 0;
  const host = new Host({ instance: {}, readProperties: () => ({}), isOpen: () => open, refresh() {} });
  host.port.prepare(() => ({ getConfiguration: () => { if (fail) throw new Error('Source refresh'); return { pages: [{ groups: [{ isGroupNameHidden: true, groupFields: [{ type: 1, targetProperty: 'values', properties: { key: 'actual', onRender() {} } }] }] }] }; }, fieldChanged() {}, configurationComplete() {}, changesApplied() {}, dispose() { disposed++; } }));
  open = true; const first = host.getConfiguration(() => ({ pages: [] })); fail = true;
  assert.equal(host.getConfiguration(() => ({ pages: [] })), first); assert.equal(disposed, 0); host.dispose(); assert.equal(disposed, 1);
});

test('unsupported source composition cannot publish a new generation or retain a failed candidate controller', () => {
  for (const group of [
    { groupName: 'Named', groupFields: [{ type: 1, properties: { key: 'real', onRender() {} } }] },
    { isGroupNameHidden: true, isCollapsed: false, groupFields: [{ type: 1, properties: { key: 'real', onRender() {} } }] },
    { isGroupNameHidden: true, groupFields: [] },
  ]) {
    const Host = hostClass(); let disposed = 0, open = false;
    const host = new Host({ instance: {}, readProperties: () => ({}), isOpen: () => open, refresh() {} });
    host.port.prepare(() => ({ getConfiguration: () => ({ pages: [{ groups: [group] }] }), fieldChanged() {}, configurationComplete() {}, changesApplied() {}, dispose() { disposed++; } }));
    open = true; assert.deepEqual(host.getConfiguration(() => ({ pages: [] })), { pages: [] }); assert.equal(disposed, 1); assert.equal(host.produced, undefined); assert.equal(host.candidate, undefined); host.dispose(); assert.equal(disposed, 1);
  }
});
