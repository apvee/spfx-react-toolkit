const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { Simulate } = require('react-dom/test-utils');
const fluent = require('@fluentui/react-components');
const { stylesFixture } = require('./fixtures/executable-examples/styles-harness.cjs');
const main = 'Explore eight style recipes with the host renderer';
function catalogue() {
  const file = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/catalogue.ts');
  const module = { exports: {} };
  new Function('module', 'exports', ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS } }).outputText)(module, module.exports);
  return module.exports.executableExamples;
}
function current(f, recipe) { f.change(f.node('scenario'), recipe); }
for (const compiled of [false, true]) {
  const boundary = compiled ? 'compiled' : 'source';
  test(`Styles ${boundary}: all eight recipes retain state and enumerate the real 43-family/220-member catalog`, t => {
    const f = stylesFixture(main, { compiled }); t.after(() => f.close());
    assert.equal(f.entry.metadata.host, 'webPart');
    f.change(f.node('layout-width'), '310.5');
    assert.match(f.rules(f.node('layout-preview')), /310\.5px/);
    const retained = f.node('layout-preview').className;
    for (const [recipe, id] of [['roles','role-preview'],['preset','preset-preview'],['catalog','catalog-preview'],['interaction','state-preview'],['container','query-a-preview'],['scrollbar','scroll-preview'],['viewport','viewport-preview']]) {
      current(f, recipe); assert.ok(f.node(id)); assert.equal(f.node('query-b'), null);
      assert.equal(f.container.querySelectorAll('[data-sx-demo=layout-preview],[data-sx-demo=role-preview],[data-sx-demo=preset-preview],[data-sx-demo=catalog-preview],[data-sx-demo=state-preview],[data-sx-demo=query-a-preview],[data-sx-demo=scroll-preview],[data-sx-demo=viewport-preview]').length, 1);
    }
    current(f, 'catalog');
    assert.equal(f.node('catalog-family').options.length, 43);
    let members = 0;
    for (const family of [...f.node('catalog-family').options].map(option => option.value)) {
      f.change(f.node('catalog-family'), family);
      for (const member of [...f.node('catalog-member').options].map(option => option.value)) {
        f.change(f.node('catalog-member'), member); assert.ok(f.rules(f.node('catalog-preview')).length, family + '.' + member); members++;
      }
    }
    assert.equal(members, 220);
    for (const [family, member, property, css, value] of [['width', 'full', 'width', 'width', '100%'], ['flex', 'row', 'display', 'display', 'flex'], ['textAlign', 'center', 'textAlign', 'text-align', 'center']]) {
      f.change(f.node('catalog-family'), family); f.change(f.node('catalog-member'), member);
      const rules = f.rules(f.node('catalog-preview'));
      assert.match(rules, new RegExp(`${css}:var\\(`));
      assert.match(rules, new RegExp(`--[^:;]*${property}[^:;]*:${value.replace('%', '%')};`));
    }
    f.change(f.node('catalog-family'), 'width'); f.change(f.node('catalog-member'), 'px'); f.change(f.node('catalog-value'), '84.5');
    f.change(f.node('catalog-family'), 'foreground'); assert.equal(f.node('catalog-value'), null);
    f.change(f.node('catalog-family'), 'width'); f.change(f.node('catalog-member'), 'px'); assert.equal(f.node('catalog-value').value, '84.5');
    f.change(f.node('catalog-family'), 'grid'); assert.equal(f.node('catalog-value').value, '1');
    current(f, 'layout'); assert.equal(f.node('layout-width').value, '310.5'); assert.equal(f.node('layout-preview').className, retained);
  });
  test(`Styles ${boundary}: reset keeps recipe, native input identity and focus while clearing an invalid baseline draft`, t => {
    const f = stylesFixture(main, { compiled, initialRevision: 17 }); t.after(() => f.close());
    const input = f.node('layout-width');
    f.change(input, ''); assert.equal(input.value, ''); assert.match(f.rules(f.node('layout-preview')), /240px/);
    f.h.act(() => input.focus());
    f.reset();
    assert.equal(f.node('scenario').value, 'layout'); assert.equal(f.node('layout-width'), input);
    assert.equal(document.activeElement, input); assert.equal(input.value, '240');
    current(f, 'interaction'); f.change(f.node('selected'), true); f.change(f.node('disabled'), true);
    const selected = f.node('selected'); f.h.act(() => selected.focus()); f.reset();
    assert.equal(f.node('scenario').value, 'interaction'); assert.equal(f.node('selected'), selected); assert.equal(document.activeElement, selected);
    assert.equal(selected.checked, false); assert.equal(f.node('disabled').checked, false); assert.equal(f.node('state-preview').disabled, false);
  });
}

test('Styles main inherits live host tokens, real contrast metadata, direction and independent renderer state', t => {
  const f = stylesFixture(main, { dir: 'rtl', theme: { ...fluent.webLightTheme, colorBrandBackground: '#ab1234' } });
  const other = stylesFixture(main, { theme: { ...fluent.webDarkTheme, colorBrandBackground: '#12abcd' } });
  t.after(() => { other.close(); f.close(); });
  current(f, 'layout'); f.change(f.node('layout-width'), '360');
  const original = f.node('layout-preview'); const className = original.className;
  f.updateTheme({ ...fluent.webLightTheme, colorBrandBackground: '#cd5678' });
  assert.equal(f.node('layout-preview'), original); assert.equal(original.className, className);
  assert.match(f.node('provider').className, /fui-FluentProvider/);
  assert.equal(f.node('provider').getAttribute('dir'), 'rtl');
  assert.equal(f.node('layout-width').value, '360'); assert.equal(other.node('layout-width').value, '240'); assert.notEqual(f.renderer, other.renderer);
  assert.equal(f.container.querySelector('[data-style-theme]').getAttribute('data-style-theme'), '#cd5678');
  assert.equal(other.container.querySelector('[data-style-theme]').getAttribute('data-style-theme'), '#12abcd');
  current(f, 'preset'); f.change(f.node('preset'), 'inverted'); f.updateTheme(fluent.teamsHighContrastTheme, 'HIGHCONTRAST');
  assert.equal(f.node('preset-preview').getAttribute('data-unsupported'), 'true'); assert.match(f.node('preset-notice').textContent, /Canvas is shown/);
  current(f, 'roles'); assert.equal(f.node('preset-notice'), null); current(f, 'preset'); assert.equal(f.node('preset').value, 'inverted');
  current(f, 'interaction'); f.change(f.node('selected'), true); f.change(f.node('disabled'), true);
  assert.equal(f.node('state-preview').getAttribute('aria-pressed'), 'true');
  assert.doesNotMatch(f.rules(f.node('state-preview')), /--apvee-sx-(?:backgroundColor|color|borderColor)-base-(?:hover|active|focus-visible):var\(/);
});

test('main reset restores every recipe baseline while keeping the selected viewport recipe', t => {
  const f = stylesFixture(main); t.after(() => f.close());
  for (const [id, value] of [['layout-width',360],['layout-gap',24],['layout-columns',3]]) f.change(f.node(id), value);
  current(f, 'roles'); f.change(f.node('foreground'), 'subtle'); f.change(f.node('background'), 'alternative');
  current(f, 'preset'); f.change(f.node('preset'), 'brand');
  current(f, 'catalog'); f.change(f.node('catalog-family'), 'padding'); f.change(f.node('catalog-member'), 'px'); f.change(f.node('catalog-value'), '33.5');
  current(f, 'interaction'); f.change(f.node('selected'), true); f.h.act(() => f.node('state-preview').click()); f.change(f.node('disabled'), true);
  f.change(f.node('direction'), 'rtl');
  current(f, 'container'); f.change(f.node('query-a-width'), '720'); f.change(f.node('responsive-threshold'), 'large');
  current(f, 'viewport'); f.change(f.node('viewport-threshold'), 'small');
  const subject = f.node('viewport-preview'); f.reset();
  assert.equal(f.node('scenario').value, 'viewport'); assert.equal(f.node('viewport-preview'), subject);
  assert.equal(f.node('viewport-threshold').value, 'medium'); assert.equal(f.node('direction').value, 'host');
  current(f, 'layout');
  for (const [id, value] of [['layout-width','240'],['layout-gap','12'],['layout-columns','2']]) assert.equal(f.node(id).value, value);
  current(f, 'roles'); assert.equal(f.node('foreground').value, 'primary'); assert.equal(f.node('background').value, 'canvas');
  current(f, 'preset'); assert.equal(f.node('preset').value, 'canvas');
  current(f, 'interaction'); assert.equal(f.node('selected').checked, false); assert.equal(f.node('disabled').checked, false); assert.match(f.container.textContent, /Action invocations: 0/);
  current(f, 'container'); assert.equal(f.node('query-a-width').value, '420'); assert.equal(f.node('responsive-threshold').value, 'medium');
  current(f, 'catalog'); assert.equal(f.node('catalog-family').value, 'width'); assert.equal(f.node('catalog-member').value, 'auto');
  f.change(f.node('catalog-member'), 'px'); assert.equal(f.node('catalog-value').value, '24');
});

test('main configuration keeps the original bounded two-column layout and narrow-screen collapse', t => {
  const f = stylesFixture(main); t.after(() => f.close());
  const controls = f.node('layout-width').closest('.fui-Field').parentElement;
  const rules = f.rules(controls);
  assert.match(rules, /grid-template-columns:repeat\(2, minmax\(0, 1fr\)\)/);
  assert.match(rules, /max-width:100%/);
  assert.match(rules, /@media.*600px/);
});

test('overflow helper labels the real nested public member paths', t => {
  const entry = catalogue().find(entry => entry.publicNames.includes('overflow') && entry.ownerKey.endsWith('#<module>'));
  const f = stylesFixture(entry.title); t.after(() => f.close());
  const select = f.container.querySelector('select');
  for (const axis of ['horizontal', 'vertical']) for (const member of ['visible', 'hidden', 'auto']) {
    const choice = `${axis}.${member}`;
    assert.ok([...select.options].some(option => option.value === choice), choice);
    f.change(select, choice); assert.match(f.container.textContent, new RegExp(`Applied: overflow\\.${axis}\\.${member}`));
    const probe = f.container.querySelector('[data-style-helper="overflow"]');
    assert.ok(probe);
    const rules = f.rules(probe), property = axis === 'horizontal' ? 'X' : 'Y';
    assert.match(rules, new RegExp(`overflow-${property.toLowerCase()}:var\\(`));
    assert.match(rules, new RegExp(`--[^:;]*overflow${property}[^:;]*:${member};`));
  }
});

for (const compiled of [false, true]) {
  test(`presets ${compiled ? 'compiled' : 'source'}: later overrides replace only their color property and reset to canvas`, t => {
    const entry = catalogue().find(entry => entry.publicNames.includes('presets') && entry.ownerKey.endsWith('#<module>'));
    const f = stylesFixture(entry.title, { compiled, initialRevision: 9 }); t.after(() => f.close());
    const probe = f.container.querySelector('[data-style-helper="presets"]');
    const preset = f.container.querySelector('select');
    const [foregroundOverride, backgroundOverride] = f.container.querySelectorAll('input[type=checkbox]');
    assert.ok(foregroundOverride, 'Real foreground override control');
    assert.ok(backgroundOverride, 'Real background override control');
    const colors = (foreground, background) => {
      const rules = f.rules(probe);
      assert.match(rules, new RegExp(`--apvee-sx-color-base:\\s*var\\(--${foreground}\\)`));
      assert.match(rules, new RegExp(`--apvee-sx-backgroundColor-base:\\s*var\\(--${background}\\)`));
    };
    assert.equal(preset.value, 'canvas'); colors('colorNeutralForeground1', 'colorNeutralBackground1');
    f.change(preset, 'brand'); colors('colorNeutralForegroundOnBrand', 'colorBrandBackground');
    f.change(foregroundOverride, true); colors('colorBrandForeground1', 'colorBrandBackground');
    f.change(backgroundOverride, true); colors('colorBrandForeground1', 'colorSubtleBackground');
    f.change(foregroundOverride, false); colors('colorNeutralForegroundOnBrand', 'colorSubtleBackground');
    f.change(backgroundOverride, false); colors('colorNeutralForegroundOnBrand', 'colorBrandBackground');
    f.change(foregroundOverride, true); f.change(backgroundOverride, true);
    f.h.act(() => foregroundOverride.focus()); f.reset();
    assert.equal(preset.value, 'canvas'); assert.equal(foregroundOverride.checked, false); assert.equal(backgroundOverride.checked, false);
    assert.equal(f.container.querySelector('[data-style-helper="presets"]'), probe);
    assert.equal(f.container.querySelector('input[type=checkbox]'), foregroundOverride); assert.equal(document.activeElement, foregroundOverride);
    colors('colorNeutralForeground1', 'colorNeutralBackground1');
  });
}

test('43 independent namespace cases and three native state cases run their real descriptors and reset without remount', () => {
  const entries = catalogue().filter(entry => entry.ownerKey.includes('/helpers/styles/') && !entry.ownerKey.includes('::'));
  assert.equal(entries.filter(entry => entry.ownerKey.endsWith('#<module>')).length, 43);
  assert.equal(entries.length, 46);
  for (const entry of entries) {
    const f = stylesFixture(entry.title);
    try {
      assert.equal(entry.metadata.host, 'none');
      const probe = f.container.querySelector('[data-style-helper]:not([data-style-helper$="-ancestor"])');
      assert.ok(probe, entry.title); assert.ok(f.rules(probe).length, entry.title);
      const descriptor = [...f.container.querySelectorAll('select')].find(select => [...select.options].some(option => option.value === 'px' || option.value === 'columns'));
      if (descriptor) {
        f.change(descriptor, [...descriptor.options].some(option => option.value === 'px') ? 'px' : 'columns');
        const input = f.container.querySelector('input[type=number]');
        if (input) { f.change(input, ''); assert.equal(input.value, ''); f.reset(); }
      } else {
        const first = f.container.querySelector('select'); if (first && first.options.length > 1) f.change(first, first.options[1].value);
        const checkbox = f.container.querySelector('input[type=checkbox]'); if (checkbox) f.change(checkbox, !checkbox.checked);
        f.reset();
      }
      assert.equal(f.container.querySelector('[data-style-helper]:not([data-style-helper$="-ancestor"])'), probe, entry.title + ' retains native subject identity');
      assert.equal(f.container.querySelector('[style]'), null, entry.title + ' authored inline style absent');
    } finally { f.close(); }
  }
});

test('native state helper cases keep selected aria, gate enabled rules and count only enabled native clicks', () => {
  const entries = catalogue().filter(entry => /\/states\.ts#(?:hover|active|focusVisible)$/.test(entry.ownerKey));
  assert.equal(entries.length, 3);
  for (const entry of entries) {
    const f = stylesFixture(entry.title);
    try {
      const action = f.container.querySelector('button');
      const [disabled, selected] = f.container.querySelectorAll('input[type=checkbox]');
      f.h.act(() => action.click()); assert.match(f.container.textContent, /Action invocations: 1/);
      f.change(selected, true); f.change(disabled, true);
      assert.equal(action.disabled, true); assert.equal(action.getAttribute('aria-pressed'), 'true');
      assert.doesNotMatch(f.rules(action), /--apvee-sx-(?:backgroundColor|color|borderColor)-base-(?:hover|active|focus-visible):var\(/);
      f.h.act(() => action.click()); assert.match(f.container.textContent, /Action invocations: 1/);
      f.reset(); assert.equal(f.container.querySelector('button'), action); assert.equal(action.disabled, false);
      assert.equal(action.getAttribute('aria-pressed'), 'false'); assert.match(f.container.textContent, /Action invocations: 0/);
    } finally { f.close(); }
  }
});

for (const compiled of [false, true]) {
  test(`focused helper layout/query contexts use actual ${compiled ? 'compiled' : 'source'} classes`, () => {
    for (const name of ['gap', 'flexItem', 'container', 'responsive', 'viewport']) {
      const entry = catalogue().find(entry => entry.publicNames.includes(name) && entry.ownerKey.endsWith('#<module>'));
      assert.ok(entry, name);
      const f = stylesFixture(entry.title, { compiled });
      try {
        const probe = f.container.querySelector(`[data-style-helper="${name}"]`);
        if (name === 'gap') {
          assert.equal(probe.children.length, 3);
          f.change(f.container.querySelector('select'), 'px');
          f.change(f.container.querySelector('input[type=number]'), '36.5');
          assert.match(f.rules(probe), /36\.5px/); assert.match(f.rules(probe), /row-gap:/); assert.match(f.rules(probe), /column-gap:/);
        } else if (name === 'flexItem') {
          assert.equal(probe.parentElement.children.length, 3);
          assert.match(f.rules(probe.parentElement), /240px/);
          f.change(f.container.querySelector('input[type=checkbox]'), true);
          assert.match(f.rules(probe.parentElement), /640px/);
          f.change(f.container.querySelector('select'), 'noGrow'); assert.match(f.rules(probe), /flex-grow:/);
        } else {
          assert.deepEqual([...f.container.querySelector('select').options].map(option => option.value), ['small','medium','large']);
          assert.match(f.rules(probe), name === 'viewport' ? /@media.*640px/ : /@container.*640px/);
          if (name === 'viewport') assert.equal(f.container.querySelector('input[type=number]'), null);
          else {
            const ancestor = f.container.querySelector(`[data-style-helper="${name}-ancestor"]`);
            f.change(f.container.querySelector('input[type=number]'), '640'); assert.match(f.rules(ancestor), /640px/);
            assert.match(f.rules(ancestor), /container-name:/);
            if (name === 'container') {
              f.change(f.container.querySelector('input[type=checkbox]'), false); assert.doesNotMatch(f.rules(ancestor), /container-name:/);
              f.reset(); assert.match(f.rules(ancestor), /container-name:/); assert.match(f.rules(ancestor), /420px/);
            }
          }
        }
      } finally { f.close(); }
    }
  });
}
