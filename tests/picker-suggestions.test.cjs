const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const load = () => createHarness().load('controls/pickers/picker-suggestions.internal.ts');
const option = (id, key = id, text = id) => ({ key, text, url: `https://tenant.test/sites/${id}`, value: { kind: 'site', url: `https://tenant.test/sites/${id}` } });
test('composition gives consumer precedence, excludes selected identities and caps all sources together', () => {
  const { composePickerSuggestions } = load();
  const sections = composePickerSuggestions({ consumer: [option('a'), option('b')], currentUser: [option('c')], people: [option('a', 'personal-a', 'Personal'), option('d')], groups: [option('e'), option('f')], selected: [option('b', 'selected-b')], limit: 5 });
  assert.deepEqual(sections.flatMap(s => s.options.map(o => o.key)), ['a', 'c', 'd', 'e', 'f']);
  assert.equal(sections[0].options[0].text, 'a');
  assert.deepEqual(composePickerSuggestions({ consumer: [option('a')], showDefaultSuggestions: false }), []);
});
test('local transformations cannot reinsert excluded identities or unverified candidates', () => {
  const { composePickerSuggestions } = load();
  const result = composePickerSuggestions({ consumer: [option('a'), option('b')], selected: [option('a')], accepts: o => o.key !== 'c', filterOptions: options => [...options, option('a'), option('c')], limit: 5 });
  assert.deepEqual(result.flatMap(s => s.options.map(o => o.key)), ['b']);
});

test('contextual Site suggestions follow consumer priority, canonical dedup and shared five-row cap', () => {
  const { composePickerSuggestions } = load();
  const consumer = option('a'), current = option('a', 'current-alias'), hub = option('hub');
  const sections = composePickerSuggestions({ consumer: [consumer], currentSite: [current], hubSite: [hub], people: Array.from({length: 8}, (_, i) => option(`extra${i}`)), limit: 5 });
  assert.deepEqual(sections.map(s => s.source), ['consumer', 'hubSite', 'people']);
  assert.equal(sections.flatMap(s => s.options).length, 5);
  const selected = composePickerSuggestions({ currentSite: [current], hubSite: [hub], selected: [consumer, hub] });
  assert.deepEqual(selected, []);
});
