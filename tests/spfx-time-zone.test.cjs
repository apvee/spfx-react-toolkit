const test = require('node:test');
const assert = require('node:assert/strict');
const Module = require('node:module');
const mapping = require('./fixtures/spfx-time-zone-mapping.json');
const { createHarness } = require('./provider-test-harness.cjs');

test('timezone audit covers every pinned candidate with explicit provenance and exclusions', () => {
  assert.equal(mapping.schemaVersion, 1);
  assert.equal(mapping.territory, '001');
  assert.equal(mapping.rows.length, 111);
  assert.equal(new Set(mapping.rows.map(row => row.sharePointId)).size, 111);
  assert.equal(mapping.rows.filter(row => row.status === 'verified').length, 104);
  assert.deepEqual(mapping.rows.filter(row => row.status === 'excluded').map(row => row.sharePointId), [7, 16, 30, 32, 41, 81, 92]);
  const sources = new Map(mapping.sources.map(source => [source.id, source]));
  for (const source of sources.values()) {
    assert.match(source.url, /^https:\/\//);
    assert.ok(source.revision);
    assert.match(source.snapshotSha256, /^[0-9a-f]{64}$/);
    if (source.license) assert.ok(source.notice.includes('Permission'));
  }
  for (const row of mapping.rows) {
    assert.ok(Number.isInteger(row.sharePointId));
    assert.ok(row.candidateWindowsId);
    assert.ok(row.candidateDescription);
    assert.ok(row.evidence.length >= 2);
    for (const sourceId of row.evidence) assert.ok(sources.has(sourceId));
    assert.ok(row.reason);
    if (row.status === 'verified') {
      assert.ok(row.windowsId);
      assert.ok(row.iana);
      assert.doesNotThrow(() => new Intl.DateTimeFormat('en-US', { timeZone: row.iana }).format(0), row.iana);
    } else {
      assert.equal(row.status, 'excluded');
      assert.equal(row.iana, undefined);
      assert.equal(row.windowsId, undefined);
    }
  }
  for (const [id, windowsId, iana] of [
    [4, 'W. Europe Standard Time', 'Europe/Berlin'],
    [10, 'Eastern Standard Time', 'America/New_York'],
    [12, 'Mountain Standard Time', 'America/Denver'],
    [13, 'Pacific Standard Time', 'America/Los_Angeles'],
    [17, 'New Zealand Standard Time', 'Pacific/Auckland'],
  ]) {
    const row = mapping.rows.find(candidate => candidate.sharePointId === id);
    assert.equal(row.windowsId, windowsId);
    assert.equal(row.iana, iana);
  }
});

function run(body) {
  const harness = createHarness();
  try { body(harness); } finally { harness.close(); }
}

function pageContext(id, culture = 'it-IT') {
  return {
    cultureInfo: { currentCultureName: culture, currentUICultureName: 'en-US', isRightToLeft: false },
    web: { timeZoneInfo: id === undefined ? undefined : { id, offset: 0, description: 'untrusted label', daylightOffset: 0, standardOffset: 0 } },
  };
}

test('private resolver follows the frozen mapping and leaves malformed or unaudited IDs unresolved', () => run(h => {
  const { resolveSPFxIanaTimeZone } = h.load('helpers/spfx-time-zone.internal.ts');
  for (const row of mapping.rows) assert.equal(resolveSPFxIanaTimeZone(row.sharePointId), row.iana);
  // null and string inputs exercise an untyped malformed boundary, not the API contract.
  for (const id of [undefined, null, '4', 4.5, -1, 0, 1, 52, 105, 113, 116, NaN, Infinity, -Infinity, Number.MAX_SAFE_INTEGER]) {
    assert.equal(resolveSPFxIanaTimeZone(id), undefined, String(id));
  }
}));

test('locale helper preserves SDK metadata and derives only from the numeric ID', () => run(h => {
  const { getSPFxLocaleInfo } = h.load('helpers/spfx-page-context.helpers.ts');
  const context = pageContext(4);
  const zone = context.web.timeZoneInfo;
  const original = { ...zone };
  Object.freeze(zone);
  const locale = getSPFxLocaleInfo(context);
  assert.equal(locale.ianaTimeZone, 'Europe/Berlin');
  assert.strictEqual(locale.timeZone, zone);
  assert.deepEqual(zone, original);
  assert.equal(locale.locale, 'it-IT');
  assert.equal(locale.uiLocale, 'en-US');
  assert.equal(locale.isRtl, false);
  const misleading = pageContext(13, 'ja-JP');
  misleading.web.timeZoneInfo.description = 'Europe/Rome';
  misleading.web.timeZoneInfo.offset = 540;
  misleading.web.timeZoneInfo.daylightOffset = 540;
  misleading.web.timeZoneInfo.standardOffset = 540;
  assert.equal(getSPFxLocaleInfo(misleading).ianaTimeZone, 'America/Los_Angeles');
  for (const id of [undefined, 7, 999]) assert.equal(getSPFxLocaleInfo(pageContext(id)).ianaTimeZone, undefined);
}));

function offset(zone, instant) {
  return new Intl.DateTimeFormat('en-US', { timeZone: zone, timeZoneName: 'shortOffset' })
    .formatToParts(new Date(instant)).find(part => part.type === 'timeZoneName').value;
}

test('representative zones retain independent winter and summer rules', () => run(h => {
  const { resolveSPFxIanaTimeZone } = h.load('helpers/spfx-time-zone.internal.ts');
  for (const [id, winter, summer] of [[4, 'GMT+1', 'GMT+2'], [13, 'GMT-8', 'GMT-7'], [17, 'GMT+13', 'GMT+12'], [12, 'GMT-7', 'GMT-6'], [38, 'GMT-7', 'GMT-7']]) {
    const zone = resolveSPFxIanaTimeZone(id);
    assert.equal(offset(zone, '2026-01-15T12:00:00Z'), winter);
    assert.equal(offset(zone, '2026-07-15T12:00:00Z'), summer);
  }
  // Eastern and old Indiana shared a winter offset, but not daylight rules in 2005.
  assert.equal(offset(resolveSPFxIanaTimeZone(10), '2005-01-15T12:00:00Z'), 'GMT-5');
  assert.equal(offset(resolveSPFxIanaTimeZone(34), '2005-01-15T12:00:00Z'), 'GMT-5');
  assert.equal(offset(resolveSPFxIanaTimeZone(10), '2005-07-15T12:00:00Z'), 'GMT-4');
  assert.equal(offset(resolveSPFxIanaTimeZone(34), '2005-07-15T12:00:00Z'), 'GMT-5');
}));

test('locale hook agrees with helper across isolated providers and in-place zone ID updates', () => run(h => {
  const pageContextKey = {};
  // The SDK runtime requires SPFx's resource loader; only its service identity is needed here.
  const originalLoad = Module._load;
  let useSPFxLocaleInfo;
  try {
    Module._load = function (specifier, ...args) {
      if (specifier === '@microsoft/sp-page-context') return { PageContext: { serviceKey: pageContextKey } };
      return originalLoad.call(this, specifier, ...args);
    };
    ({ useSPFxLocaleInfo } = h.load('hooks/useSPFxLocaleInfo.ts'));
  } finally { Module._load = originalLoad; }
  const { getSPFxLocaleInfo } = h.load('helpers/spfx-page-context.helpers.ts');
  function observer(id, context) {
    const host = h.instance(id);
    const consume = host.context.serviceScope.consume;
    host.context.serviceScope.consume = key => key === pageContextKey ? context : consume(key);
    let latest;
    function Child() { latest = useSPFxLocaleInfo(); return h.React.createElement(h.React.Fragment); }
    const fixture = h.mount(host, Child);
    return { fixture, get latest() { return latest; } };
  }
  const contextA = pageContext(4);
  const contextB = pageContext(13, 'en-US');
  const a = observer('a', contextA);
  const b = observer('b', contextB);
  assert.deepEqual(a.latest, getSPFxLocaleInfo(contextA));
  assert.equal(a.latest.ianaTimeZone, 'Europe/Berlin');
  assert.equal(b.latest.ianaTimeZone, 'America/Los_Angeles');
  assert.strictEqual(a.latest.timeZone, contextA.web.timeZoneInfo);
  const originalZone = contextA.web.timeZoneInfo;
  contextA.web.timeZoneInfo.id = 17;
  a.fixture.render();
  assert.equal(a.latest.ianaTimeZone, 'Pacific/Auckland');
  assert.strictEqual(a.latest.timeZone, originalZone);
  assert.deepEqual(a.latest, getSPFxLocaleInfo(contextA));
  assert.equal(b.latest.ianaTimeZone, 'America/Los_Angeles');
  const stable = a.latest;
  a.fixture.render();
  assert.strictEqual(a.latest, stable);
  contextA.web.timeZoneInfo.id = 7;
  a.fixture.render();
  assert.equal(a.latest.ianaTimeZone, undefined);
  assert.equal(b.latest.ianaTimeZone, 'America/Los_Angeles');
  a.fixture.unmount();
  b.fixture.render();
  assert.deepEqual(b.latest, getSPFxLocaleInfo(contextB));
}));
