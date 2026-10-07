const assert = require('node:assert/strict');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { test } = require('node:test');

function runScenario(scenario) {
  const result = spawnSync(process.execPath, [
    path.join(__dirname, 'fixtures/tree-shaking/runtime/pnp-standalone.cjs'), scenario
  ], { encoding: 'utf8', timeout: 30000 });
  assert.equal(result.error, undefined);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  const report = JSON.parse(result.stdout.trim());
  assert.equal(report.scenario, scenario);
}

// Removing the list leaf's webs registration must break real query construction.
test('list factory works without context factory or webs preload', () => runScenario('list'));

// Missing context webs/batching, search, or generic batching imports break these requests.
test('context factory registers web queries and batching without other factories', () => runScenario('context'));
test('search factory queries and suggests without context or list registrations', () => runScenario('search'));
test('generic service batches real requests with only caller-owned web registration', () => runScenario('batch'));
