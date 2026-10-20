const assert = require('node:assert/strict');
/** Shared acceptance for the same actual app/panels against source and compiled controls. */
async function accessTrimsBrowser({ page, check, select, singleExample, labels, overflow, capture }) {
  const configure = patch => page.evaluate(patch => window.trimFixture.configure(patch), patch);
  const release = () => page.evaluate(() => window.trimFixture.release());
  const state = async expected => { await page.locator(`[data-trim-demo=state][data-trim-state=${expected}]`).waitFor(); check(`observed ${expected} branch`, await page.locator('[data-trim-demo=protected]').count(), 0); };
  const allowed = async () => { await page.locator('[data-trim-demo=protected]').waitFor(); check('allowed protected mount', await page.locator('[data-trim-demo=state]').count(), 0); };
  for (const key of ['sharepoint-permission-trim', 'sharepoint-group-trim', 'entra-group-trim']) {
    await select(page, key); await singleExample(page, key); await labels(page, key);
    check(`${key}: no check before Run`, await page.locator('[data-trim-demo=idle]').count(), 1);
    await configure({ mode: 'allowed', hold: true });
    await page.locator('[data-trim-demo=run]').click(); await state('loading');
    check(`${key}: no protected effects while loading`, await page.locator('[data-trim-demo=mount-gauge]').textContent(), 'Protected mounts: 0; active: false');
    await release(); await allowed(); await page.locator('[data-trim-demo=increment]').click();
    check(`${key}: counter active`, await page.locator('[data-trim-demo=increment]').textContent(), 'Protected counter: 1');
    const previousTarget = await page.locator('[data-trim-demo=applied-target]').textContent();
    const previousReads = await page.evaluate(() => window.trimFixture.calls.length);
    if (key === 'sharepoint-permission-trim') await page.locator('[data-trim-demo=web-url]').fill('https://simulated.invalid/sites/next');
    else await page.locator('[data-trim-demo=groups]').fill('Changed group');
    check(`${key}: draft keeps applied identity`, await page.locator('[data-trim-demo=applied-target]').textContent(), previousTarget);
    check(`${key}: draft keeps counter`, await page.locator('[data-trim-demo=increment]').textContent(), 'Protected counter: 1');
    check(`${key}: draft issues no reads`, await page.evaluate(() => window.trimFixture.calls.length), previousReads);
    await configure({ mode: 'denied', hold: true }); await page.locator('[data-trim-demo=run]').click(); await state('loading');
    await release(); await state('denied');
    check(`${key}: target actually changed`, (await page.locator('[data-trim-demo=applied-target]').textContent()) !== previousTarget, true);
    check(`${key}: inactive after denial`, (await page.locator('[data-trim-demo=mount-gauge]').textContent()).includes('active: false'), true);
    await configure({ mode: 'error', hold: true }); await page.locator('[data-trim-demo=run]').click(); await state('loading'); await release(); await state('error');
    await configure({ mode: 'allowed', hold: true }); await page.locator('[data-trim-demo=run]').click(); await state('loading'); await release(); await allowed();
    check(`${key}: recovery mounts fresh counter`, await page.locator('[data-trim-demo=increment]').textContent(), 'Protected counter: 0');
    check(`${key}: exactly two positive mounts`, (await page.locator('[data-trim-demo=mount-gauge]').textContent()).includes('Protected mounts: 2; active: true'), true);
    if (key === 'entra-group-trim') {
      await page.locator('[data-trim-demo=group-type]').selectOption('M365');
      await page.locator('[data-trim-demo=relationship]').selectOption('owner');
      await page.locator('[data-trim-demo=match]').selectOption('any');
      await page.locator('[data-trim-demo=run]').click(); await allowed();
      await page.waitForFunction(() => document.querySelector('[data-trim-demo=applied-target]')?.textContent?.includes('"relationship":"owner"'));
      check('Entra owner reads native direct owners', await page.evaluate(() => window.trimFixture.calls.some(call => call.url.includes('/owners'))), true);
      await page.locator('[data-trim-demo=relationship]').selectOption('memberOrOwner'); await page.locator('[data-trim-demo=run]').click(); await allowed();
      check('Entra relationship/match applied', (await page.locator('[data-trim-demo=applied-target]').textContent()).includes('"match":"any"'), true);
      await page.locator('[data-trim-demo=group-type]').selectOption('Security'); await page.locator('[data-trim-demo=run]').click(); await state('error');
      await page.locator('[data-trim-demo=group-type]').selectOption('M365'); await page.locator('[data-trim-demo=run]').click(); await allowed();
    }
    for (const width of [1280, 320]) { await page.setViewportSize({ width, height: 1000 }); await overflow(page, `${key}/${width}`); await capture(page, `${key}-${width}`); }
    await page.setViewportSize({ width: 1280, height: 1000 });
  }
  await select(page, 'sharepoint-permission-trim');
  await page.locator('[data-demo-variant=sharepoint-permission-trim]').selectOption('and');
  await configure({ mode: 'denied', hold: false });
  const beforeAND = await page.evaluate(() => window.trimFixture.calls.length);
  await page.locator('[data-trim-demo=run]').click(); await state('denied');
  check('AND denied outer never starts inner group read', await page.evaluate(index => window.trimFixture.calls.slice(index).some(call => call.transport === 'graph' && call.url.startsWith('/groups')), beforeAND), false);
  await configure({ mode: 'allowed', hold: false }); await page.locator('[data-trim-demo=run]').click(); await allowed();
  check('AND allowed starts actual inner group read', await page.evaluate(index => window.trimFixture.calls.slice(index).some(call => call.transport === 'graph' && call.url.startsWith('/groups')), beforeAND), true);
  await select(page, 'context'); await page.locator('[data-demo-variant=context]').selectOption('scoped');
  await configure({ mode: 'allowed', hold: true }); await page.locator('[data-context-demo=run-scoped]').click();
  await page.waitForFunction(() => document.querySelector('[data-context-demo=scoped-state]')?.textContent === 'Loading scoped permissions');
  await release(); await page.waitForFunction(() => document.querySelector('[data-context-demo=scoped-state]')?.textContent === 'Allowed: view list items');
  for (const scope of ['list', 'item', 'folder']) {
    await page.locator('[data-trim-demo=scope]').selectOption(scope);
    if (scope === 'folder') { await page.locator('[data-trim-demo=list-value]').fill('Shared Documents'); await page.locator('[data-trim-demo=folder-path]').fill('/sites/demo/Shared Documents/Reports'); }
    await page.locator('[data-context-demo=run-scoped]').click();
    await page.waitForFunction(() => document.querySelector('[data-context-demo=scoped-state]')?.textContent === 'Allowed: view list items');
    check(`Context ${scope} applied target`, (await page.locator('[data-context-demo=scoped-target]').textContent()).includes(`"scope":"${scope}"`), true);
  }
  await configure({ mode: 'error', hold: false }); await page.locator('[data-context-demo=run-scoped]').click();
  await page.waitForFunction(() => document.querySelector('[data-context-demo=scoped-state]')?.textContent?.startsWith('Error:'));
  check('Context actual scoped error distinct from denial', (await page.locator('[data-context-demo=scoped-state]').textContent()).startsWith('Error:'), true);
  await configure({ mode: 'allowed', hold: false });
  const calls = await page.evaluate(() => window.trimFixture.calls);
  assert.ok(calls.some(call => call.transport === 'sp')); assert.ok(calls.some(call => call.transport === 'graph' && call.url === '/me/checkMemberGroups'));
  check('Native read-only boundaries: no fixture CRUD', calls.every(call => !call.body || call.url === '/me/checkMemberGroups'), true);
}
module.exports = { accessTrimsBrowser };
