const assert = require('node:assert/strict');
const test = require('node:test');
const { pilotTitles, filterPilotGroups } = require('./fixtures/demo-presentation/jsdoc-pilots.cjs');

test('browser pilot filter selects exactly four features and nine independently named scenarios', () => {
  const features = Object.entries(pilotTitles).map(([title, titles]) => ({ key: title, title, scenarios: [...titles.map(title => ({ title })), { title: 'Future scenario' }] }));
  const groups = [{ key: 'controls', features: features.slice(0, 2) }, { key: 'hooks', features: features.slice(2) }, { key: 'future', features: [{ title: 'Future feature', scenarios: [] }] }];
  const filtered = filterPilotGroups(groups);
  assert.equal(filtered.length, 2);
  assert.equal(filtered.flatMap(group => group.features).length, 4);
  assert.equal(filtered.flatMap(group => group.features).flatMap(feature => feature.scenarios).length, 9);
  assert.equal(groups[0].features[0].scenarios.length, 4, 'filter never mutates the app catalogue');
  assert.throws(() => filterPilotGroups(groups.slice(0, 1)), /Missing pilot feature/);
  assert.throws(() => filterPilotGroups([{ ...groups[0], features: [{ ...features[0], scenarios: [] }, features[1]] }, groups[1]]), /Missing pilot scenario/);
});
