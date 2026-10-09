const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createRequire } = require('node:module');
const harness = require('./fixtures/formatted-text-harness.cjs');
const root = path.resolve(__dirname, '..');
// The pre-install probe runs unchanged against a checkout-independent candidate install.
const dependencyRoot = process.env.FORM_DEPENDENCY_PROBE_ROOT || root;
const hostRequire = createRequire(path.join(dependencyRoot, 'package.json'));
const runtimeRequire = dependencyRoot === root ? createRequire(path.join(root, 'packages/spfx-react-toolkit/package.json')) : hostRequire;
const baseline = require('./fixtures/dependency-baseline.json');
const versions = { ...baseline.fixed, ...baseline.shared };
function packageInfo(name) {
  let directory = path.dirname(runtimeRequire.resolve(name.startsWith('@types/') ? `${name}/package.json` : name));
  while (!fs.existsSync(path.join(directory, 'package.json')) ||
      JSON.parse(fs.readFileSync(path.join(directory, 'package.json'))).name !== name) {
    const parent = path.dirname(directory);
    assert.notEqual(parent, directory, `Package metadata missing for ${name}`);
    directory = parent;
  }
  return JSON.parse(fs.readFileSync(path.join(directory, 'package.json')));
}
test('exact dependency baseline preserves the fixed host and supported React peers', () => {
  const semver = require('semver');
  for (const [name, version] of Object.entries(versions)) {
    const info = packageInfo(name);
    assert.equal(info.version, version, name);
    if (info.peerDependencies?.react) assert.ok(semver.satisfies('17.0.1', info.peerDependencies.react), name);
  }
  assert.ok(semver.satisfies(versions['react-hook-form'], packageInfo('@hookform/resolvers').peerDependencies['react-hook-form']));
  for (const compat of ['@fluentui/react-datepicker-compat', '@fluentui/react-timepicker-compat', '@fluentui/react-calendar-compat']) {
    const compatRequire = createRequire(hostRequire.resolve(compat));
    for (const name of ['react', 'react-dom', '@griffel/react', '@fluentui/react-shared-contexts', '@fluentui/react-utilities']) {
      assert.equal(fs.realpathSync(compatRequire.resolve(name)), fs.realpathSync(hostRequire.resolve(name)), `${compat}: ${name}`);
    }
  }
  if (dependencyRoot === root) {
    for (const workspace of ['packages/spfx-react-toolkit', 'apps/spfx-react-toolkit-test']) {
      const workspaceRequire = createRequire(path.join(root, workspace, 'package.json'));
      for (const name of ['react', 'react-dom', 'react-hook-form', '@fluentui/react-components', '@fluentui/react-shared-contexts', '@griffel/react']) {
        assert.equal(fs.realpathSync(workspaceRequire.resolve(name)), fs.realpathSync(hostRequire.resolve(name)), `${workspace}: ${name}`);
      }
    }
    const library = require('../packages/spfx-react-toolkit/package.json');
    for (const name of ['react-hook-form', '@fluentui/react-datepicker-compat', '@fluentui/react-timepicker-compat', '@fluentui/react-icons']) {
      assert.equal(library.devDependencies[name], versions[name]);
      assert.equal(library.peerDependencies[name], `>=${versions[name]} <${versions[name].split('.')[0] * 1 + 1}.0.0`);
      assert.equal(library.dependencies[name], undefined);
    }
    assert.equal(library.dependencies['@hookform/resolvers'], versions['@hookform/resolvers']);
    assert.equal(library.dependencies['@standard-schema/spec'], versions['@standard-schema/spec']);
    assert.equal(library.dependencies.zod, undefined); assert.equal(library.dependencies.yup, undefined);
  }
});
test('React17/TS5.3 infer distinct input/output, controller, array and compat callback types', () => {
  const ts = hostRequire('typescript');
  const fixture = dependencyRoot === root ? path.join(__dirname, 'fixtures/form-dependency-types.tsx') : path.join(dependencyRoot, 'types.tsx');
  const program = ts.createProgram([fixture], {
    noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020,
    module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs,
    jsx: ts.JsxEmit.React, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'],
  });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: file => file, getCurrentDirectory: () => dependencyRoot, getNewLine: () => '\n',
  }));
});
test('actual React17/Fluent/RHF submit, reset and date/time callbacks preserve live values', async () => {
  const dom = harness.createSxDom();
  const browserGlobals = ['NodeFilter', 'Node', 'HTMLElement', 'Element', 'Event', 'CustomEvent'];
  const originalGlobals = new Map(browserGlobals.map(name => [name, global[name]]));
  for (const name of browserGlobals) global[name] = dom.window[name];
  // JSDOM lacks layout observation; callback semantics do not depend on geometry.
  dom.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  Object.defineProperty(dom.window, 'crypto', { value: require('node:crypto').webcrypto });
  const React = hostRequire('react');
  const ReactDOM = hostRequire('react-dom');
  const { act, Simulate } = hostRequire('react-dom/test-utils');
  const { FluentProvider, webLightTheme, Input } = hostRequire('@fluentui/react-components');
  const { DatePicker } = hostRequire('@fluentui/react-datepicker-compat');
  const { TimePicker } = hostRequire('@fluentui/react-timepicker-compat');
  const { useForm, useController, FormProvider, useFormContext } = hostRequire('react-hook-form');
  const { standardSchemaResolver } = runtimeRequire('@hookform/resolvers/standard-schema');
  const submissions = [], dates = [], times = [];
  let methods, context;
  const initialDate = new Date(2026, 7, 30, 12);
  function TextControl() {
    context = useFormContext();
    const { field } = useController({ name: 'amount', control: context.control });
    return React.createElement(Input, { id: 'amount', value: field.value, onChange: (_, data) => field.onChange(data.value), ref: field.ref });
  }
  function Probe() {
    methods = useForm({ defaultValues: { amount: '3', date: initialDate, time: undefined }, resolver: standardSchemaResolver({
      '~standard': { version: 1, vendor: 'probe', validate: async value => ({ value: { ...value, amount: Number(value.amount) } }) }
    }) });
    return React.createElement(FormProvider, methods, React.createElement('form', { onSubmit: methods.handleSubmit(value => submissions.push(value)) },
      React.createElement(TextControl),
      React.createElement(DatePicker, { id: 'date', allowTextInput: true, value: methods.watch('date'), formatDate: date => String(date.getDate()),
        parseDateFromString: text => new Date(2026, 7, Number(text), 12), onSelectDate: date => { dates.push(date); methods.setValue('date', date); } }),
      React.createElement(TimePicker, { id: 'time', freeform: true, dateAnchor: new Date(2026, 7, 30), hourCycle: 'h23',
        onTimeChange: (_, data) => { times.push(data); methods.setValue('time', data.selectedTime); } })));
  }
  const container = document.createElement('div'); document.body.appendChild(container);
  try {
    await act(async () => { ReactDOM.render(React.createElement(FluentProvider, { theme: webLightTheme }, React.createElement(Probe)), container); });
    assert.equal(context.control, methods.control);
    await act(async () => { Simulate.change(container.querySelector('#amount'), { target: { value: '8' } }); });
    assert.equal(methods.getValues('amount'), '8');
    const date = container.querySelector('#date');
    await act(async () => { Simulate.change(date, { target: { value: '31' } }); });
    await act(async () => { Simulate.blur(date); });
    assert.equal(dates.at(-1).getDate(), 31);
    const time = container.querySelector('#time');
    await act(async () => { Simulate.change(time, { target: { value: '14:30' } }); });
    await act(async () => { Simulate.keyDown(time, { key: 'Enter', code: 'Enter', which: 13 }); });
    assert.equal(times.at(-1).selectedTime.getHours(), 14); assert.equal(times.at(-1).selectedTime.getMinutes(), 30);
    await act(async () => { Simulate.submit(container.querySelector('form')); });
    assert.equal(submissions.length, 1); assert.equal(submissions[0].amount, 8);
    assert.equal(submissions[0].date.getDate(), 31); assert.equal(submissions[0].time.getHours(), 14);
    await act(async () => { methods.reset(); });
    assert.equal(container.querySelector('#amount').value, '3');
    assert.equal(methods.getValues('date').getDate(), 30); assert.equal(methods.getValues('time'), undefined);
  } finally {
    await act(async () => { ReactDOM.unmountComponentAtNode(container); });
    // Let native Tabster effects finish before closing their owning window.
    await new Promise(resolve => setTimeout(resolve, 0));
    container.remove();
    dom.window.close();
    for (const [name, value] of originalGlobals) { if (value === undefined) delete global[name]; else global[name] = value; }
  }
});

test('opt-in webpack boundary retains the host RHF implementation', { skip: dependencyRoot !== root }, async () => {
  try {
    const loaded = await harness.loadModule('tests/fixtures/form-dependency-types.tsx', { reactHookForm: true });
    for (const name of ['useForm', 'useController', 'useFieldArray']) {
      assert.equal(loaded.dependencyIdentity[name], hostRequire('react-hook-form')[name]);
    }
  } finally { await harness.cleanup(); }
});
