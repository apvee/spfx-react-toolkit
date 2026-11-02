const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const panelPath = path.resolve(__dirname, 'fixtures/foundation-legacy/StylesPanel.tsx');
function fixture(t, hostTheme) {
  assert.ok(fs.existsSync(panelPath), 'The observable Styles panel is missing');
  const toolkit = {};
  let choice = 'host', updateHost;
  const fluent = require('@fluentui/react-components');
  // Only unavailable SPFx context boundaries are doubled. React, descriptors,
  // useSx, Griffel, FluentProvider and the theme adapter are the real modules.
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit,
    './useSPFxTeams': { useSPFxTeams: () => ({ supported: choice !== 'host', theme: choice === 'host' ? undefined : choice }) },
    './useSPFxThemeInfo': { useSPFxThemeInfo: () => hostTheme } });
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'),
    h.load('hooks/useSPFxFluent9ThemeInfo.ts'), h.load('helpers/spfx-theme.helpers.ts'));
  const renderer = createDOMRenderer(null);
  const Panel = h.load(panelPath).default;
  const view = h.mountComponent(() => {
    const [theme, setTheme] = h.React.useState(toolkit.createFluent9ThemeFromSPFxTheme(hostTheme));
    updateHost = next => { choice = next; setTheme(toolkit.getTeamsFluentTheme(next)); };
    return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme }, h.React.createElement(Panel)));
  });
  t.after(h.close);
  const node = id => { const e = view.container.querySelector(`[data-sx-demo="${id}"]`); assert.ok(e, id); return e; };
  const change = (id, value) => h.act(() => {
    if (id === 'theme') { updateHost(value); return; }
    const e = node(id);
    if (e.type === 'checkbox') { e.checked = value; Simulate.change(e, { target: e }); }
    else { e.value = String(value); Simulate.change(e, { target: e }); }
  });
  const rules = id => {
    const classes = new Set(node(id).className.split(/\s+/));
    return Object.keys(renderer.insertionCache).filter(rule => [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(m => classes.has(m[1]))).join('\n');
  };
  return { ...view, node, change, rules, toolkit, h };
}
test('Styles controls update real width, gap and grid descriptors', t => {
  const f = fixture(t);
  f.change('layout-width', 360); f.change('layout-gap', 24); f.change('layout-columns', 3);
  assert.match(f.rules('layout-preview'), /360px/);
  assert.match(f.rules('layout-preview'), /24px/);
  assert.match(f.rules('layout-preview'), /repeat\(3, minmax\(0, 1fr\)\)/);
  f.change('layout-width', 240);
  assert.match(f.rules('layout-preview'), /240px/);
  assert.doesNotMatch(f.rules('layout-preview'), /360px/);
});
test('Styles host mode carries real SPFx neutral state colors into the shared FluentProvider', t => {
  const host = require('@fluentui/react').createTheme();
  const f = fixture(t, host);
  assert.equal(f.container.querySelector('[data-sx-demo=theme]'), null);
  const css = window.getComputedStyle(f.node('provider'));
  assert.equal(css.getPropertyValue('--colorNeutralStrokeAccessible').trim(), '#605e5c');
  assert.equal(css.getPropertyValue('--colorNeutralStrokeAccessibleHover').trim(), '#323130');
  assert.equal(css.getPropertyValue('--colorNeutralStrokeAccessiblePressed').trim(), '#201f1e');
});
test('Styles foreground, subtle/alternative, preset and real theme provider controls are observable', t => {
  const f = fixture(t);
  f.change('foreground', 'subtle');
  assert.match(f.rules('role-preview'), /var\(--colorNeutralForeground2\)/);
  f.change('background', 'subtle');
  assert.match(f.rules('role-preview'), /var\(--colorSubtleBackground\)/);
  f.change('background', 'alternative');
  assert.match(f.rules('role-preview'), /var\(--colorNeutralBackground2\)/);
  f.change('preset', 'brand');
  const classes = f.node('preset-preview').className;
  f.change('theme', 'dark');
  assert.equal(f.node('preset-preview').className, classes);
  const themeRules = Array.from(document.querySelectorAll('style')).flatMap(tag => tag.sheet ? Array.from(tag.sheet.cssRules) : []).map(rule => rule.cssText).join('\n');
  assert.ok(themeRules.includes(`--colorBrandBackground: ${f.toolkit.getTeamsFluentTheme('dark').colorBrandBackground}`));
  f.change('theme', 'contrast'); f.change('preset', 'inverted');
  assert.ok(f.node('preset-notice').textContent.includes('unsupported'));
  assert.equal(f.node('preset-preview').hasAttribute('data-unsupported'), true);
});
test('disabled Fluent action removes enabled state recipes and selected presentation', t => {
  const f = fixture(t);
  assert.match(f.rules('state-preview'), /:hover/);
  assert.match(f.rules('state-preview'), /:active/);
  assert.match(f.rules('state-preview'), /:focus-visible/);
  f.change('selected', true); f.change('disabled', true);
  assert.equal(f.node('state-preview').disabled, true);
  assert.equal(f.node('state-preview').getAttribute('aria-pressed'), 'true');
  assert.match(f.rules('state-preview'), /var\(--colorNeutralBackgroundDisabled\)/);
  // Fluent retains its own focus/forced-colors rules; descriptor state assignments are absent.
  assert.doesNotMatch(f.rules('state-preview'), /--apvee-sx-(?:backgroundColor|color|borderColor)-base-(?:hover|active|focus-visible):var\(/);
  f.change('disabled', false);
  assert.match(f.rules('state-preview'), /:hover/);
});
test('every catalog member applies real classes and independent query/scrollbar regions exist', t => {
  const f = fixture(t);
  const family = f.node('catalog-family');
  assert.ok(family.options.length >= 43);
  for (const option of Array.from(family.options)) {
    f.change('catalog-family', option.value);
    for (const member of Array.from(f.node('catalog-member').options)) {
      f.change('catalog-member', member.value);
      assert.ok(f.node('catalog-preview').className, `${option.value}.${member.value}`);
      assert.ok(f.rules('catalog-preview').length, `${option.value}.${member.value} emits actual Griffel rules`);
    }
  }
  assert.match(f.rules('query-a'), /container-name:\s*var\(/);
  assert.match(f.rules('query-b'), /container-name:\s*var\(/);
  f.change('query-a-width', 720); f.change('query-b-width', 420);
  assert.match(f.rules('query-a'), /720px/); assert.match(f.rules('query-b'), /420px/);
  assert.match(f.rules('scroll-preview'), /--apvee-sx-scrollbarWidth-base:\s*thin/);
  assert.match(f.rules('scroll-preview'), /overflow-y:/);
});
test('panel has no inline style assignments and fractional pixels remain observable', t => {
  const f = fixture(t);
  assert.doesNotMatch(fs.readFileSync(panelPath, 'utf8'), /\bstyle\s*=/);
  f.change('layout-width', 240.5); f.change('layout-gap', 24.5);
  assert.match(f.rules('layout-preview'), /240\.5px/);
  assert.match(f.rules('layout-preview'), /24\.5px/);
  f.change('catalog-family', 'padding'); f.change('catalog-member', 'px'); f.change('catalog-value', 24.5);
  assert.match(f.rules('catalog-preview'), /24\.5px/);
  f.change('catalog-family', 'grid');
  assert.equal(f.node('catalog-value').value, '1');
  f.change('catalog-value', 2.5);
  assert.equal(f.node('catalog-value').value, '2.5');
  f.h.act(() => Simulate.blur(f.node('catalog-value')));
  assert.equal(f.node('catalog-value').value, '1');
  assert.match(f.rules('catalog-preview'), /repeat\(1, minmax\(0, 1fr\)\)/);
});

test('numeric drafts allow clearing and sequential replacement while retaining valid descriptors', t => {
  const f = fixture(t);
  f.change('layout-width', '');
  assert.equal(f.node('layout-width').value, '');
  assert.match(f.rules('layout-preview'), /240px/);
  for (const value of ['3', '32', '320', '320.5']) f.change('layout-width', value);
  assert.equal(f.node('layout-width').value, '320.5');
  assert.match(f.rules('layout-preview'), /320\.5px/);
  f.change('layout-columns', ''); f.change('layout-columns', '0');
  assert.match(f.rules('layout-preview'), /repeat\(2, minmax\(0, 1fr\)\)/);
  f.h.act(() => Simulate.blur(f.node('layout-columns')));
  assert.equal(f.node('layout-columns').value, '2');
});

test('hook, provider, style and service panels use Fluent configuration and shared demo regions', () => {
  const componentRoot = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
  const panels = ['Providers', 'Search'];
  const files = panels.map(name => path.join(componentRoot, 'panels', name + 'Panel.tsx'))
    .concat(fs.readdirSync(path.join(componentRoot, 'demos')).filter(name => /^PnP.*\.tsx$/.test(name)).map(name => path.join(componentRoot, 'demos', name)));
  for (const file of files) {
    const source = fs.readFileSync(file, 'utf8');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]/, path.basename(file));
    assert.doesNotMatch(source, /<(?:input|select|textarea|button|details|a)\b/, path.basename(file));
    assert.doesNotMatch(source, /\bstyles?\s*=/, path.basename(file));
    const delegated = /from ['"]\.\.\/demos['"]/.test(source) && /<PnP[A-Za-z]+Demo/.test(source);
    assert.ok(/DemoCard|useDemoStyles/.test(source) || delegated, `${path.basename(file)} must own shared regions or delegate an actual demo component`);
  }
  const generatedRoot = path.join(componentRoot, 'generated/jsdoc-examples');
  const catalogueModule = { exports: {} };
  const catalogueText = fs.readFileSync(path.join(generatedRoot, 'catalogue.ts'), 'utf8');
  new Function('module', 'exports', require('typescript').transpileModule(catalogueText, {
    compilerOptions: { module: require('typescript').ModuleKind.CommonJS },
  }).outputText)(catalogueModule, catalogueModule.exports);
  const crudCases = catalogueModule.exports.executableExamples.filter(entry => entry.ownerKey === 'packages/spfx-react-toolkit/src/hooks/useSPFxPnPList.ts#useSPFxPnPList' && entry.title === 'Manage list items with paging and captured delete consent');
  assert.equal(crudCases.length, 1, 'canonical CRUD replaces the final legacy PnP panel');
  const crudSource = fs.readFileSync(path.join(generatedRoot, crudCases[0].id + '.tsx'), 'utf8');
  assert.deepEqual([crudCases[0].metadata.host, crudCases[0].metadata.kind, crudCases[0].metadata.contentPlacement], ['webPart', 'write', 'configuration']);
  assert.doesNotMatch(crudSource, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|summary|a)\b|<FluentProvider\b|\bstyle\s*=/);
  assert.match(crudSource, /className=/, 'portable CRUD authors classes with the existing host renderer');
  const configurationTitles = {
    useSPFxConfigurationStore: 'Compare layered configuration and physical inventories',
    createSPFxConfigurationStoreService: 'Compose tenant defaults and site overrides with the service',
  };
  const configurationOwners = ['packages/spfx-react-toolkit/src/hooks/useSPFxConfigurationStore.ts#useSPFxConfigurationStore', 'packages/spfx-react-toolkit/src/services/spfx-configuration-store.service.ts#createSPFxConfigurationStoreService'];
  const configurationCases = catalogueModule.exports.executableExamples.filter(entry => configurationOwners.includes(entry.ownerKey));
  assert.deepEqual(configurationCases.map(entry => [entry.featureTitle, entry.title]).sort(), Object.entries(configurationTitles).sort());
  for (const entry of configurationCases) {
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.deepEqual([entry.metadata.host, entry.metadata.kind, entry.metadata.contentPlacement], ['webPart', 'write', 'configuration']);
    assert.equal(entry.groupKey, entry.featureTitle === 'useSPFxConfigurationStore' ? 'hooks' : 'services');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|summary|a)\b|<FluentProvider\b/);
    const parser = require('typescript'); const parsed = parser.createSourceFile(entry.id + '.tsx', source, parser.ScriptTarget.Latest, true, parser.ScriptKind.TSX);
    let calls = 0;
    function inspectConfiguration(node) {
      if (parser.isCallExpression(node) && parser.isIdentifier(node.expression) && node.expression.text === entry.featureTitle) calls++;
      if (parser.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text), 'Configuration authored styles use classes');
      parser.forEachChild(node, inspectConfiguration);
    }
    inspectConfiguration(parsed); assert.equal(calls, 1, 'each case invokes its own actual public owner once');
    assert.match(source, /return render \? render\(regions\)/);
    assert.match(source, /overflowX: 'auto'/);
  }
  const tenantTitles = {
    useSPFxTenantKeyValueStore: 'Explore tenant records and exact text persistence',
    createSPFxTenantKeyValueStoreService: 'Round trip numeric looking text through a tenant service',
  };
  const tenantOwners = ['packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore.ts#useSPFxTenantKeyValueStore', 'packages/spfx-react-toolkit/src/services/spfx-tenant-key-value-store.service.ts#createSPFxTenantKeyValueStoreService'];
  const tenantCases = catalogueModule.exports.executableExamples.filter(entry => tenantOwners.includes(entry.ownerKey));
  assert.deepEqual(tenantCases.map(entry => [entry.featureTitle, entry.title]).sort(), Object.entries(tenantTitles).sort(), 'both actual source-owned Tenant list cases replace the retired panel');
  for (const entry of tenantCases) {
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.deepEqual([entry.metadata.host, entry.metadata.kind, entry.metadata.contentPlacement], ['webPart', 'write', 'configuration']);
    assert.equal(entry.groupKey, entry.featureTitle === 'useSPFxTenantKeyValueStore' ? 'hooks' : 'services');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|summary|a)\b|<FluentProvider\b/);
    const parser = require('typescript'); const parsed = parser.createSourceFile(entry.id + '.tsx', source, parser.ScriptTarget.Latest, true, parser.ScriptKind.TSX);
    let calls = 0;
    function inspectTenant(node) {
      if (parser.isCallExpression(node) && parser.isIdentifier(node.expression) && node.expression.text === entry.featureTitle) calls++;
      if (parser.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text), 'Tenant authored styles use classes');
      parser.forEachChild(node, inspectTenant);
    }
    inspectTenant(parsed); assert.equal(calls, 1, 'each case invokes its own actual public owner once');
    assert.match(source, /return render \? render\(regions\)/, 'standalone and shared renderers retain one controller');
    assert.match(source, /overflowX: 'auto'/);
  }
  // A control using a helper does not become a selectable hook scenario.
  const hooks = catalogueModule.exports.executableExamples
    .filter(entry => entry.groupKey === 'hooks' && entry.metadata.host === 'none')
    .map(entry => fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8'));
  assert.equal(hooks.length, 3, 'two callback pilots and the complete import-interoperability scenario own shared configuration regions');
  const runtimeOwners = [
    'packages/spfx-react-toolkit/src/hooks/useSPFxProperties.ts#useSPFxProperties',
    'packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts#useSPFxSessionStorage',
    'packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts#useSPFxLocalStorage',
    'packages/spfx-react-toolkit/src/hooks/useSPFxLogger.ts#useSPFxLogger',
    'packages/spfx-react-toolkit/src/hooks/useSPFxPerformance.ts#useSPFxPerformance',
  ];
  const runtimeCases = catalogueModule.exports.executableExamples.filter(entry => runtimeOwners.includes(entry.ownerKey));
  assert.deepEqual(runtimeCases.map(entry => entry.ownerKey).sort(), [...runtimeOwners].sort(), 'five actual Runtime owners replace the retired panel');
  for (const entry of runtimeCases) {
    assert.equal(entry.metadata.host, 'webPart', entry.title);
    assert.equal(entry.metadata.contentPlacement, 'configuration', entry.title);
  }
  for (const source of hooks) {
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|a)\b|\bstyles?\s*=/);
    assert.match(source, /return render\(\{/);
    assert.match(source, /contentPlacement: 'configuration'/);
  }
  const ts = require('typescript');
  const clientTitles = {
    useSPFxAadTokenProvider: 'Inspect client readiness and token-provider initialization',
    useSPFxHttpClient: 'Load an external todo with tracked invocation state',
    useSPFxSPHttpClient: 'Read the current SharePoint web title and URL',
    useSPFxMSGraphClient: 'Read the current Graph profile after initialization',
    useSPFxAadHttpClient: 'Initialize a custom API client and call a path',
    useSPFxApiPermissionPrecheck: 'Manually check Graph and custom API delegated scopes',
  };
  const clientOwners = Object.keys(clientTitles).map(name => `packages/spfx-react-toolkit/src/hooks/${name}.ts#${name}`);
  const clientCases = catalogueModule.exports.executableExamples.filter(entry => clientOwners.includes(entry.ownerKey));
  assert.deepEqual(clientCases.map(entry => [entry.featureTitle, entry.title]).sort(), Object.entries(clientTitles).sort(), 'exact six canonical Clients source cases transfer the retired panel checks');
  for (const entry of clientCases) {
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.equal(entry.metadata.host, 'webPart', entry.title);
    assert.equal(entry.metadata.contentPlacement, 'configuration', entry.title);
    assert.equal(entry.metadata.kind, entry.featureTitle === 'useSPFxAadTokenProvider' ? 'snapshot' : 'read', entry.title);
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|summary|a)\b/, entry.title);
    const parsed = ts.createSourceFile(entry.id + '.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    function inspectClient(node) {
      if (ts.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text), `${entry.title} uses authored classes`);
      ts.forEachChild(node, inspectClient);
    }
    inspectClient(parsed);
    assert.match(source, new RegExp(`\\b${entry.featureTitle}\\(`), 'case calls its actual public subject');
    assert.match(source, /return render \? render\(regions\)/, 'case forwards actual shared regions with a portable fallback');
    assert.doesNotMatch(source, /<FluentProvider\b/, 'the host owns the renderer and theme');
    if (entry.metadata.kind === 'read') {
      assert.match(source, /commands:\s*<>\s*<Button\b/, 'explicit real Fluent commands survive migration');
      assert.match(source, /makeStyles\(\{ diagnostics: \{ minWidth: 0, maxWidth: '100%', overflowX: 'auto'/, 'diagnostics retain bounded authored CSS');
      assert.match(source, /<Accordion\b/, 'raw diagnostics use actual Fluent accordion controls');
    } else assert.doesNotMatch(source, /<(?:Button|Input|Select)\b/, 'token readiness has no invented token command');
    if (['useSPFxAadHttpClient', 'useSPFxApiPermissionPrecheck'].includes(entry.featureTitle)) assert.match(source, /controls:[\s\S]*<Field\b[\s\S]*<Input\b/, 'resource configuration retains real Fluent fields');
  }
  const graphTitles = {
    useSPFxUserPhoto: 'Read the current user photo with managed URL cleanup',
    useSPFxOneDriveAppData: 'Read and explicitly replace OneDrive app data',
  };
  const graphOwners = Object.keys(graphTitles).map(name => `packages/spfx-react-toolkit/src/hooks/${name}.ts#${name}`);
  const graphCases = catalogueModule.exports.executableExamples.filter(entry => graphOwners.includes(entry.ownerKey));
  assert.deepEqual(graphCases.map(entry => [entry.featureTitle, entry.title]).sort(), Object.entries(graphTitles).sort(), 'exact two canonical GraphData cases replace the panel');
  for (const entry of graphCases) {
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.equal(entry.metadata.host, 'webPart'); assert.equal(entry.metadata.contentPlacement, 'configuration');
    assert.equal(entry.metadata.kind, entry.featureTitle === 'useSPFxUserPhoto' ? 'read' : 'write');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|summary|a)\b|<FluentProvider\b/);
    const parsed = ts.createSourceFile(entry.id + '.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    let subjectCalls = 0;
    function inspectGraph(node) {
      if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === entry.featureTitle) subjectCalls++;
      if (ts.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text), entry.title + ' uses classes');
      ts.forEachChild(node, inspectGraph);
    }
    inspectGraph(parsed);
    assert.equal(subjectCalls, 1, 'case calls its actual public subject, including generic invocations');
    assert.match(source, /return render \? render\(regions\)/);
    assert.match(source, /commands:\s*<>\s*<Button\b/);
    if (entry.featureTitle === 'useSPFxOneDriveAppData') {
      assert.match(source, /<Field\b[\s\S]*<Select\b/); assert.match(source, /<Input\b/);
      assert.match(source, /<Accordion\b/); assert.match(source, /overflowX: 'auto'/);
    } else assert.match(source, /<Image\b[\s\S]*<Avatar\b/);
  }
  const siteTitles = {
    useSPFxSiteKeyValueStore: 'Explore collection root records with guarded mutations',
    createSPFxSiteKeyValueStoreService: 'Read a collection root record before a consented save',
  };
  const siteOwners = ['packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore.ts#useSPFxSiteKeyValueStore', 'packages/spfx-react-toolkit/src/services/spfx-site-key-value-store.service.ts#createSPFxSiteKeyValueStoreService'];
  const siteCases = catalogueModule.exports.executableExamples.filter(entry => siteOwners.includes(entry.ownerKey));
  assert.deepEqual(siteCases.map(entry => [entry.featureTitle, entry.title]).sort(), Object.entries(siteTitles).sort());
  for (const entry of siteCases) {
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.deepEqual([entry.metadata.host, entry.metadata.kind, entry.metadata.contentPlacement], ['webPart', 'write', 'configuration']);
    assert.equal(entry.groupKey, entry.featureTitle === 'useSPFxSiteKeyValueStore' ? 'hooks' : 'services');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|summary|a)\b|<FluentProvider\b/);
    const parsed = ts.createSourceFile(entry.id + '.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    let calls = 0;
    function inspectSite(node) {
      if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === entry.featureTitle) calls++;
      if (ts.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text), 'Site authored styles are classes');
      ts.forEachChild(node, inspectSite);
    }
    inspectSite(parsed); assert.equal(calls, 1);
    assert.match(source, /<Field\b[\s\S]*<Input\b/); assert.match(source, /<Checkbox\b/); assert.match(source, /overflowX: 'auto'/);
  }
  const propertyOwners = {
    useSPFxTenantProperty: ['packages/spfx-react-toolkit/src/hooks/useSPFxTenantProperty.ts#useSPFxTenantProperty', 'Read tenant version and counter with explicit activation'],
    createSPFxTenantPropertyService: ['packages/spfx-react-toolkit/src/services/spfx-tenant-property.service.ts#createSPFxTenantPropertyService', 'Read and validate a catalog StorageEntity'],
  };
  const propertyCases = catalogueModule.exports.executableExamples.filter(entry => Object.values(propertyOwners).some(([owner]) => owner === entry.ownerKey));
  assert.equal(propertyCases.length, 2, 'exact two independent tenant-property owners');
  for (const entry of propertyCases) {
    assert.deepEqual([entry.ownerKey, entry.title], propertyOwners[entry.featureTitle]);
    assert.deepEqual([entry.metadata.host, entry.metadata.kind, entry.metadata.contentPlacement], ['webPart', 'read', 'configuration']);
    assert.equal(entry.groupKey, entry.featureTitle === 'useSPFxTenantProperty' ? 'hooks' : 'services');
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|summary|a)\b|<FluentProvider\b/);
    const parsed = ts.createSourceFile(entry.id + '.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    let calls = 0;
    function inspectProperty(node) {
      if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === entry.featureTitle) calls++;
      if (ts.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text));
      ts.forEachChild(node, inspectProperty);
    }
    inspectProperty(parsed); assert.equal(calls, entry.featureTitle === 'useSPFxTenantProperty' ? 2 : 1);
    assert.match(source, /overflowX: 'auto'/);
  }
  // Context's finite acceptance cohort must not absorb later client/data hook families.
  const contextNames = ['Context', 'PageContext', 'InstanceInfo', 'DisplayMode', 'IsEdit', 'EnvironmentInfo',
    'UserInfo', 'SiteInfo', 'ListInfo', 'LocaleInfo', 'HubSiteInfo', 'Teams', 'PageType', 'CorrelationInfo',
    'Permissions', 'CrossSitePermissions', 'ContainerInfo', 'ContainerSize', 'ThemeInfo', 'Fluent9ThemeInfo', 'ServiceScope']
    .map(name => 'useSPFx' + name);
  const expectedContext = Object.entries(require('./fixtures/demo-presentation/jsdoc-pilots.cjs').pilotTitles)
    .filter(([name]) => contextNames.includes(name));
  const contextTitles = expectedContext.flatMap(([, titles]) => titles);
  const contextCases = catalogueModule.exports.executableExamples.filter(entry => contextTitles.includes(entry.title));
  assert.equal(expectedContext.length, 21, 'Context acceptance includes 21 actual APIs');
  assert.equal(contextTitles.length, 22, 'the two cross-site overloads retain distinct cases');
  assert.deepEqual(contextCases.map(entry => entry.title).sort(), [...contextTitles].sort(), 'all actual Context cases transfer the retired panel checks');
  for (const entry of contextCases) {
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.equal(entry.metadata.host, 'webPart', entry.title);
    assert.equal(entry.metadata.contentPlacement, 'configuration', entry.title);
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|a)\b/, entry.title);
    const parsed = ts.createSourceFile(entry.id + '.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    function inspect(node) {
      if (ts.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text), `${entry.title} uses host classes`);
      ts.forEachChild(node, inspect);
    }
    inspect(parsed);
    assert.match(source, /return render \? render\(regions\)/, 'source case forwards its actual regions');
    if (entry.metadata.kind === 'read') assert.match(source, /commands:\s*<Button\b/, 'Hub and both cross-site cases expose actual explicit Fluent commands');
    else if (entry.featureTitle === 'useSPFxLocaleInfo') assert.match(source, /controls:\s*<Field\b[\s\S]*<Select\b/, 'Locale has its real seasonal choice');
    else assert.doesNotMatch(source, /<(?:Button|Input|Select)\b/, 'readonly native snapshots require no invented controls');
  }
  for (const entry of runtimeCases) {
    const source = fs.readFileSync(path.join(generatedRoot, entry.id + '.tsx'), 'utf8');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]|<(?:input|select|textarea|button|details|a)\b/, entry.title);
    // Class maps named styles are valid; reject actual JSX style attributes.
    const parsed = ts.createSourceFile(entry.id + '.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    function inspect(node) {
      if (ts.isJsxAttribute(node)) assert.ok(!['style', 'styles'].includes(node.name.text), `${entry.title} must use classes rather than inline style attributes`);
      ts.forEachChild(node, inspect);
    }
    inspect(parsed);
    assert.match(source, /commands:\s*<>\s*<Button\b/, `${entry.title} exposes actual Fluent commands`);
    if (['useSPFxProperties', 'useSPFxLocalStorage', 'useSPFxSessionStorage'].includes(entry.featureTitle)) {
      assert.match(source, /controls:\s*<(?:Field|Input)\b/, `${entry.title} exposes actual Fluent input configuration`);
    }
    assert.match(source, /return render\(\{/);
    assert.match(source, /contentPlacement: 'configuration'/);
  }
  const stylesCase = catalogueModule.exports.executableExamples.find(entry => entry.title === 'Explore eight style recipes with the host renderer');
  assert.ok(stylesCase);
  const stylesSource = fs.readFileSync(path.join(generatedRoot, stylesCase.id + '.tsx'), 'utf8');
  assert.doesNotMatch(stylesSource, /from ['\"]@fluentui\/react['\"]|<(?:input|select|textarea|button|details|a)\b|\bstyles?\s*=/);
  assert.match(stylesSource, /return render\(\{/);
  assert.match(stylesSource, /contentPlacement: 'configuration'/);
  assert.match(fs.readFileSync(path.join(componentRoot, 'shared/ExecutableExampleHost.tsx'), 'utf8'), /<DemoCard/, 'host owns the shared presentation for source examples');
});

test('Styles native slots keep their Fluent label associations', t => {
  const f = fixture(t);
  for (const id of ['layout-width', 'layout-columns', 'catalog-family', 'disabled', 'selected']) {
    const control = f.node(id);
    assert.ok(control.id, id + ' has a native control ID');
    assert.ok(Array.from(f.container.querySelectorAll('label')).some(label => label.htmlFor === control.id), id + ' is associated with its visible label');
  }
});

test('Styles bounds layout, catalog and query subjects in distinct focusable named regions', t => {
  const f = fixture(t);
  for (const [id, name] of [['layout-scroll', 'Layout preview'], ['catalog-scroll', 'Catalog preview'], ['query-scroll', 'Container query previews']]) {
    const region = f.node(id);
    assert.equal(region.getAttribute('role'), 'region');
    assert.equal(region.getAttribute('aria-label'), name);
    assert.equal(region.tabIndex, 0);
  }
  assert.equal(f.node('catalog-preview').getAttribute('role'), null, 'inner subject remains free to demonstrate its selected descriptor');
});
