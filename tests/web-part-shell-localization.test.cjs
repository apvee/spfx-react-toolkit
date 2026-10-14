const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const locales = require('./fixtures/sharepoint-ui-locales.json');

// Reviewed individual lexical overlaps, not a language-wide permission to copy English.
// Each entry binds culture, exact field and exact spelling; the final column records why.
const intentionalSharedTerms = [
  ['ca-ES', 'normal', 'Normal', 'Catalan adjective for ordinary formatting'],
  ['cy-GB', 'colorNames.brown', 'Brown', 'Established Welsh colour adjective brown'],
  ['da-DK', 'normal', 'Normal', 'Danish adjective for ordinary formatting'],
  ['da-DK', 'colorNames.orange', 'Orange', 'Danish colour name orange'],
  ['da-DK', 'colorNames.pink', 'Pink', 'Established Danish colour name pink'],
  ['de-DE', 'colorNames.orange', 'Orange', 'German colour name orange'],
  ['es-ES', 'normal', 'Normal', 'Spanish adjective for ordinary formatting'],
  ['fr-FR', 'normal', 'Normal', 'French adjective for ordinary formatting'],
  ['fr-FR', 'colorNames.orange', 'Orange', 'French colour name orange'],
  ['gl-ES', 'normal', 'Normal', 'Galician adjective for ordinary formatting'],
  ['id-ID', 'normal', 'Normal', 'Established Indonesian adjective normal'],
  ['ms-MY', 'normal', 'Normal', 'Established Malay adjective normal'],
  ['nb-NO', 'normal', 'Normal', 'Norwegian adjective for ordinary formatting'],
  ['pt-BR', 'normal', 'Normal', 'Brazilian Portuguese adjective for ordinary formatting'],
  ['pt-PT', 'normal', 'Normal', 'European Portuguese adjective for ordinary formatting'],
  ['ro-RO', 'normal', 'Normal', 'Romanian adjective for ordinary formatting'],
  ['sv-SE', 'normal', 'Normal', 'Swedish adjective for ordinary formatting'],
  ['sv-SE', 'colorNames.orange', 'Orange', 'Swedish colour name orange'],
  ['tr-TR', 'normal', 'Normal', 'Established Turkish adjective normal'],
  ['da-DK', 'colorNames.cyan', 'Cyan', 'Danish printing colour term cyan; https://lex.dk/cyan'],
  ['de-DE', 'colorNames.cyan', 'Cyan', 'German printing colour name Cyan; https://www.hp.com/ch-de/shop/products/supplies/hp-730-designjet-druckerpatrone-cyan-300-ml-p2v68a'],
  ['fr-FR', 'colorNames.cyan', 'Cyan', 'French colour noun cyan; https://www.dictionnaire-academie.fr/article/A9C5382'],
  ['nb-NO', 'colorNames.cyan', 'Cyan', 'Bokmål colour noun cyan; https://naob.no/ordbok/cyan'],
  ['sv-SE', 'colorNames.cyan', 'Cyan', 'Swedish printing colour name cyan; https://support.hp.com/se-sv/document/c03255765'],
];
function scalarLabels(strings) {
  return [
    ...Object.entries(strings).filter(([, value]) => typeof value === 'string'),
    ...Object.entries(strings.colorNames).map(([key, value]) => [`colorNames.${key}`, value]),
  ];
}
function assertTranslatedLabels(culture, strings, english) {
  const colorNames = Object.values(strings.colorNames).map(value => value.trim().normalize('NFC').toLowerCase());
  assert.equal(new Set(colorNames).size, colorNames.length, `${culture}: duplicate color name`);
  if (culture === 'en-US') return;
  const englishLabels = new Map(scalarLabels(english));
  for (const [key, value] of scalarLabels(strings)) {
    if (value !== englishLabels.get(key)) continue;
    assert.ok(intentionalSharedTerms.some(([allowedCulture, allowedKey, allowedValue]) =>
      allowedCulture === culture && allowedKey === key && allowedValue === value),
    `${culture}.${key}: unapproved English label ${JSON.stringify(value)}`);
  }
}
function fixture(t) {
  const h = createHarness();
  t.after(() => h.close());
  return h.load('controls/web-part-shell/localization/web-part-shell-locales.internal.ts');
}
test('all 50 supported UI cultures have complete distinct catalogs and typed interpolation', t => {
  const api = fixture(t);
  assert.equal(locales.cultures.length, 50);
  assert.deepEqual(Object.keys(api.webPartShellLocaleCatalogs).sort(), [...locales.cultures].sort());
  const english = api.resolveWebPartShellStrings('en-US');
  for (const culture of locales.cultures) {
    const strings = api.resolveWebPartShellStrings(culture);
    assertTranslatedLabels(culture, strings, english);
    assert.equal(strings, api.webPartShellLocaleCatalogs[culture]);
    assert.deepEqual(Object.keys(strings).sort(), Object.keys(english).sort(), culture);
    for (const [key, value] of Object.entries(strings)) {
      if (typeof value === 'string') {
        assert.ok(value.trim(), `${culture}.${key}`);
        assert.doesNotMatch(value, /TODO|PLACEHOLDER|\{(?:level|value|label|shortcut)\}/, `${culture}.${key}`);
      }
    }
    assert.deepEqual(Object.keys(strings.colorNames).sort(), Object.keys(english.colorNames).sort(), culture);
    for (const name of Object.values(strings.colorNames)) assert.ok(name.trim(), culture);
    for (const [key, args, expected] of [
      ['headingLevel', [4], ['4']], ['lineSpacingValue', [1.4], ['1.4']],
      ['formattingShortcut', ['AUTHOR_LABEL', 'Ctrl/Cmd+B'], ['AUTHOR_LABEL', 'Ctrl/Cmd+B']],
      ['mixedFormatting', ['AUTHOR_LABEL'], ['AUTHOR_LABEL']]
    ]) {
      const value = strings[key](...args);
      assert.equal(typeof value, 'string');
      for (const token of expected) assert.ok(value.includes(token), `${culture}.${key}: ${value}`);
      assert.doesNotMatch(value, /\{(?:level|value|label|shortcut)\}/);
    }
    if (culture !== 'en-US') {
      assert.notEqual(strings, english, culture);
      assert.notEqual(strings.titleEditorLoadError, english.titleEditorLoadError, culture);
      assert.notEqual(strings.savedTitleFormattingWarning, english.savedTitleFormattingWarning, culture);
    }
  }
});
test('UI culture resolver normalizes case and separators and permits only explicit aliases', t => {
  const api = fixture(t);
  for (const [input, culture] of [['IT_it', 'it-IT'], [' fr-fr ', 'fr-FR'], ['en-GB', 'en-US'], ['de-AT', 'de-DE'],
    ['zh-Hans', 'zh-CN'], ['zh-SG', 'zh-CN'], ['zh-Hant', 'zh-TW'], ['zh-HK', 'zh-TW'],
    ['sr-Latn', 'sr-Latn-RS'], ['sr-Cyrl', 'sr-Cyrl-RS'], ['pt-BR', 'pt-BR'], ['pt-PT', 'pt-PT'],
    ['az-Latn', 'az-Latn-AZ'], ['bs-Latn', 'bs-Latn-BA'], ['nb', 'nb-NO']]) {
    assert.equal(api.resolveWebPartShellStrings(input), api.webPartShellLocaleCatalogs[culture], input);
  }
  for (const input of [undefined, '', 'xx-XX', 'pt', 'zh', 'sr', 'az-Cyrl-AZ', 'sr-Arab-RS', 'pt-AO', 'zh-JP', 'constructor', '__proto__', 42, {}]) {
    assert.equal(api.resolveWebPartShellStrings(input), api.webPartShellLocaleCatalogs['en-US'], String(input));
  }
});
test('different instances resolve independently without changing existing resources', t => {
  const api = fixture(t);
  const italian = api.resolveWebPartShellStrings('it-IT');
  const original = italian.titlePlaceholder;
  const arabic = api.resolveWebPartShellStrings('ar-SA');
  const english = api.resolveWebPartShellStrings('en-US');
  assert.notEqual(italian.titlePlaceholder, arabic.titlePlaceholder);
  assert.notEqual(english.titlePlaceholder, arabic.titlePlaceholder);
  assert.equal(italian.titlePlaceholder, original);
  assert.ok(Object.isFrozen(italian));
  assert.ok(Object.isFrozen(italian.colorNames));
});
test('formatters preserve literal inputs without interpreting replacement syntax or embedded placeholders', t => {
  const api = fixture(t);
  for (const culture of locales.cultures) {
    const strings = api.resolveWebPartShellStrings(culture);
    const label = '$& {shortcut}';
    const shortcut = '$` $\' {label}';
    const formatted = strings.formattingShortcut(label, shortcut);
    assert.ok(formatted.includes(label), `${culture}: label`);
    assert.ok(formatted.includes(shortcut), `${culture}: shortcut`);
    assert.ok(strings.mixedFormatting(label).includes(label), `${culture}: mixed`);
  }
});
test('each locale resource declares all interpolation tokens exactly once', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const root = path.join(__dirname, '../packages/spfx-react-toolkit/src/controls/web-part-shell/localization/locales');
  for (const culture of locales.cultures) {
    const source = fs.readFileSync(path.join(root, `${culture}.ts`), 'utf8');
    for (const [key, tokens] of [
      ['headingLevelTemplate', ['level']], ['lineSpacingValueTemplate', ['value']],
      ['formattingShortcutTemplate', ['label', 'shortcut']], ['mixedFormattingTemplate', ['label']]
    ]) {
      const match = source.match(new RegExp(`\\b${key}: ("(?:[^"\\\\]|\\\\.)*"),`));
      assert.ok(match, `${culture}.${key}`);
      const template = JSON.parse(match[1]);
      assert.deepEqual([...template.matchAll(/\{([^}]+)\}/g)].map(item => item[1]).sort(), tokens.sort(), `${culture}.${key}`);
    }
  }
});

test('translation validator rejects copied English UI and colour labels and duplicate colour names', t => {
  const api = fixture(t);
  const italian = api.resolveWebPartShellStrings('it-IT');
  const english = api.resolveWebPartShellStrings('en-US');
  assert.throws(() => assertTranslatedLabels('it-IT', { ...italian, bold: english.bold }, english),
    /it-IT\.bold: unapproved English label "Bold"/);
  assert.throws(() => assertTranslatedLabels('it-IT', {
    ...italian, colorNames: { ...italian.colorNames, red: english.colorNames.red },
  }, english), /it-IT\.colorNames\.red: unapproved English label "Red"/);
  assert.throws(() => assertTranslatedLabels('it-IT', {
    ...italian, colorNames: { ...italian.colorNames, red: italian.colorNames.orange },
  }, english), /it-IT: duplicate color name/);
  // A legitimate shared term in one culture/field cannot permit English in another.
  assert.throws(() => assertTranslatedLabels('it-IT', { ...italian, normal: english.normal }, english),
    /it-IT\.normal: unapproved English label "Normal"/);
  assert.throws(() => assertTranslatedLabels('da-DK', {
    ...api.resolveWebPartShellStrings('da-DK'), colorNames: { ...api.resolveWebPartShellStrings('da-DK').colorNames, brown: 'Brown' },
  }, english), /da-DK\.colorNames\.brown: unapproved English label "Brown"/);
});
test('intentional shared spelling approvals are exact, justified and still present in their catalogs', t => {
  const api = fixture(t);
  const englishLabels = new Map(scalarLabels(api.resolveWebPartShellStrings('en-US')));
  const seen = new Set();
  for (const [culture, key, value, reason] of intentionalSharedTerms) {
    assert.ok(locales.cultures.includes(culture), culture);
    assert.notEqual(culture, 'en-US');
    assert.equal(typeof reason, 'string');
    assert.ok(reason.trim(), `${culture}.${key}: missing lexical justification`);
    assert.equal(englishLabels.get(key), value, `${culture}.${key}: stale English spelling`);
    assert.equal(new Map(scalarLabels(api.resolveWebPartShellStrings(culture))).get(key), value,
      `${culture}.${key}: stale local spelling`);
    const identity = JSON.stringify([culture, key, value]);
    assert.ok(!seen.has(identity), `duplicate approval ${identity}`);
    seen.add(identity);
  }
});
