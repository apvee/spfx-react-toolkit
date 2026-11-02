const { pilotTitles } = require('../tests/fixtures/demo-presentation/jsdoc-pilots.cjs');

async function checkLegacyPermissionResult(panel, check) {
  const result = panel.getByLabel('Legacy permission result', { exact: true });
  await result.getByText('Site mask: Available; Web mask: Available; List mask: Unavailable', { exact: true }).waitFor();
  check('Context legacy loading absent', await result.getByText('Loading legacy permissions', { exact: true }).count(), 0);
  check('Context legacy errors absent', await result.getByRole('alert').count(), 0);
  check('Context legacy decision labels', await result.locator('dt').allTextContents(),
    ['Manage Web', 'Manage Lists', 'Add List Items', 'Site view list items', 'List view list items']);
  check('Context legacy decision values', await result.locator('dd').allTextContents(),
    ['false', 'false', 'true', 'true', 'false']);
}

// Local browser acceptance uses real generated modules/controls/Shell. Only SPFx
// native services and HTTP transport are doubled by the integrated fixture.
async function jsdocPilotsBrowser({ page, check, settle, overflow, capture }) {
  const navigation = page.getByRole('navigation', { name: 'Gruppi e funzionalità' });
  const group = name => navigation.getByRole('button', { name: new RegExp(`^${name}`) }).first();
  const selectedPanel = title => page.getByRole('region', { name: title, exact: true });
  const reset = () => page.getByRole('button', { name: 'Ripristina scenario', exact: true });
  async function feature(title) {
    const family = ['WebPartShell', 'Pagination'].includes(title) ? 'Controlli' : ['createSPFxSiteKeyValueStoreService', 'createSPFxTenantPropertyService', 'createSPFxTenantKeyValueStoreService'].includes(title) ? 'Servizi' : 'Hook';
    if (await group(family).getAttribute('aria-expanded') !== 'true') await group(family).click();
    await navigation.getByRole('button', { name: title, exact: true }).click();
    await selectedPanel(title).locator('[data-demo-section]').waitFor();
    await settle(page);
    return selectedPanel(title);
  }
  async function scenario(title, name) {
    const panel = await feature(title);
    const selector = panel.getByRole('combobox', { name: 'Scenario', exact: true });
    if (pilotTitles[title].length > 1) {
      await selector.focus();
      await selector.selectOption({ label: name });
      await panel.getByRole('heading', { name, exact: true }).waitFor();
      check(`${name}: scenario selector keeps focus`, await selector.evaluate(element => element === document.activeElement), true);
    }
    await panel.locator('[data-demo-section]').waitFor(); await settle(page);
    check(`${name}: one mounted scenario`, await page.locator('[data-demo-section]').count(), 1);
    check(`${name}: observation available`, (await panel.textContent()).includes('Cosa osservare:'), true);
    check(`${name}: diagnostics initially closed`, await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).getAttribute('aria-expanded'), 'false');
    return panel;
  }
  async function restore(panel, name) {
    await reset().click(); await settle(page);
    check(`${name}: reset focus`, await reset().evaluate(element => element === document.activeElement), true);
    check(`${name}: reset keeps one scenario`, await panel.locator('[data-demo-section]').count(), 1);
  }
  const text = async (panel, fragment) => check(fragment, (await panel.textContent()).includes(fragment), true);
  const rows = panel => panel.locator('tbody tr').evaluateAll(nodes => nodes.map(node => Number(node.getAttribute('data-order-id'))));
  const calls = () => page.evaluate(() => window.pnpReadFixture.calls.length);
  check('Initial fixture mount issues no business requests', await calls(), 0);
  check('JSDoc filter: accepted features', await navigation.locator('[data-example-feature]').count(), Object.keys(pilotTitles).length);
  check('JSDoc filter: controls, hooks and services groups', await navigation.locator('button[aria-expanded]').count(), 3);
  let panel = await feature('WebPartShell');
  await group('Hook').click();
  check('Opening hooks leaves controls open', await group('Controlli').getAttribute('aria-expanded'), 'true');
  const heading = await panel.getByRole('heading', { level: 2 }).elementHandle();
  await group('Controlli').click();
  check('Closed group identifies active selection', (await group('Controlli').textContent()).includes('WebPartShell'), true);
  check('Closing group preserves scenario DOM', await heading.evaluate(element => element.isConnected), true); await heading.dispose();
  check('Independent hooks remain open', await group('Hook').getAttribute('aria-expanded'), 'true');
  check('Group expansion and collapse issue no business requests', await calls(), 0);

  panel = await scenario('WebPartShell', pilotTitles.WebPartShell[0]);
  await panel.getByRole('combobox', { name: 'Appearance', exact: true }).selectOption('filled');
  await panel.getByRole('combobox', { name: 'Card size', exact: true }).selectOption('small');
  await panel.getByRole('combobox', { name: 'Title position', exact: true }).selectOption('inside');
  await panel.getByRole('combobox', { name: 'Spacing', exact: true }).selectOption('large');
  await panel.getByRole('button', { name: 'Show title', exact: true }).click();
  await panel.getByRole('button', { name: 'Hide title', exact: true }).click();
  check('Hiding title focuses show command', await panel.getByRole('button', { name: 'Show title', exact: true }).evaluate(element => element === document.activeElement), true);
  await text(panel, 'Mode: Read');
  for (const name of ['Plain title', 'Formatted title', 'Empty title']) check(`Read mode disables ${name}`, await panel.getByRole('button', { name, exact: true }).isDisabled(), true);
  await restore(panel, 'Shell configuration');
  await panel.getByRole('button', { name: 'Show title', exact: true }).click();
  await panel.getByRole('combobox', { name: 'Surface', exact: true }).selectOption('application');
  check('Source surface configures actual outer Shell', await page.locator('[data-web-part-shell-surface]').getAttribute('data-web-part-shell-surface'), 'application');
  await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click();
  await panel.getByRole('combobox', { name: 'Direction', exact: true }).selectOption('rtl');
  check('Source direction configures actual outer Shell', await page.locator('[data-web-part-shell-surface]').evaluate(node => getComputedStyle(node).direction), 'rtl');
  await panel.getByRole('combobox', { name: 'Direction', exact: true }).selectOption('host');
  await page.evaluate(() => window.shellExampleFixture.setMode('Edit')); await settle(page);
  const editor = page.locator('[data-web-part-shell-header] [contenteditable=true]'); await editor.waitFor();
  const editorIdentity = await editor.elementHandle();
  await panel.getByRole('switch', { name: 'Custom title naming', exact: true }).check();
  check('Source naming updates actual title editor', await editor.getAttribute('aria-label'), 'Project section heading');
  await panel.getByRole('switch', { name: 'Custom editor labels', exact: true }).check();
  await editor.focus();
  await page.getByRole('toolbar', { name: 'Project heading controls', exact: true }).waitFor();
  check('Source editor dictionary updates actual toolbar', await page.getByRole('button', { name: 'Strong', exact: true }).count(), 1);
  check('Independent dictionaries retain editor instance', await editor.evaluate((node, original) => node === original, editorIdentity), true);
  await panel.getByRole('switch', { name: 'Custom title naming', exact: true }).uncheck();
  check('Releasing naming preserves editor command dictionary', await editor.getAttribute('aria-label') !== 'Project section heading', true);
  await editor.focus(); await page.getByRole('toolbar', { name: 'Project heading controls', exact: true }).waitFor();
  await panel.getByRole('switch', { name: 'Custom editor labels', exact: true }).uncheck();
  await panel.getByRole('combobox', { name: 'Header action', exact: true }).selectOption('ordinary');
  check('Source shorthand preserves action DOM', await page.locator('[data-shell-demo=action-slot]').count(), 0);
  await panel.getByRole('combobox', { name: 'Header action', exact: true }).selectOption('slot');
  check('Source callback creates actual action div', await page.locator('[data-shell-demo=action-slot] > [data-shell-demo=action]').count(), 1);
  await page.locator('[data-shell-demo=action]').click();
  check('Source callback action toggles actual title properties', await page.locator('[data-web-part-shell-header]').count(), 0);
  await panel.getByRole('button', { name: 'Show title', exact: true }).click();
  await page.evaluate(() => window.shellExampleFixture.setMode('Read')); await settle(page); await editorIdentity.dispose();
  await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click();
  await capture(page, 'jsdoc-shell-configuration');

  panel = await scenario('WebPartShell', pilotTitles.WebPartShell[1]);
  await panel.getByRole('button', { name: 'Use Card baseline', exact: true }).click(); await settle(page);
  const note = panel.getByRole('textbox', { name: 'Project note', exact: true });
  await note.fill('Retained content'); await panel.getByRole('button', { name: 'Increment counter', exact: true }).click();
  const retained = await note.elementHandle();
  for (const [name, value] of [['Card appearance', 'filled'], ['Card appearance', 'filled-alternative'], ['Card size', 'large'], ['Title position', 'inside']]) {
    await panel.getByRole('combobox', { name, exact: true }).selectOption(value); await settle(page);
    check(`Shell ${name}/${value} retains note DOM`, await retained.evaluate(element => element.isConnected && element.value === 'Retained content'), true);
    await text(panel, 'Child counter: 1');
  }
  await page.locator('[data-shell-demo=action]').click(); await settle(page);
  check('Host title toggle retains note DOM', await retained.evaluate(element => element.isConnected), true);
  await retained.dispose(); await restore(panel, 'Shell retained content'); await capture(page, 'jsdoc-shell-retention');

  panel = await scenario('WebPartShell', pilotTitles.WebPartShell[2]);
  await panel.getByRole('textbox', { name: 'Project note', exact: true }).fill('Must remount');
  await panel.getByRole('button', { name: 'Increment counter', exact: true }).click();
  const previous = await panel.getByRole('textbox', { name: 'Project note', exact: true }).elementHandle();
  await panel.getByRole('button', { name: 'Use subtle baseline', exact: true }).click(); await settle(page);
  // Ensure a known subtle starting point before observing the real Card transition.
  await panel.getByRole('textbox', { name: 'Project note', exact: true }).fill('Actual subtle note');
  const subtleNote = await panel.getByRole('textbox', { name: 'Project note', exact: true }).elementHandle();
  await panel.getByRole('button', { name: 'Use Card baseline', exact: true }).click(); await settle(page);
  check('Actual Shell remount detaches former child', await subtleNote.evaluate(element => element.isConnected), false);
  check('Shell remount preserves explorer scenario', await panel.getByRole('combobox', { name: 'Scenario', exact: true }).locator('option:checked').textContent(), pilotTitles.WebPartShell[2]);
  check('Shell remount resets child note', await panel.getByRole('textbox', { name: 'Project note', exact: true }).inputValue(), 'Intranet project review');
  await text(panel, 'Child counter: 0'); await previous.dispose(); await subtleNote.dispose(); await capture(page, 'jsdoc-shell-remount');

  panel = await scenario('Pagination', pilotTitles.Pagination[0]);
  check('Known rows baseline', await rows(panel), [1, 2, 3, 4, 5]);
  await panel.getByRole('button', { name: 'Next page', exact: true }).click();
  check('Known rows next page', await rows(panel), [6, 7, 8, 9, 10]);
  await restore(panel, 'Known rows'); check('Known rows reset', await rows(panel), [1, 2, 3, 4, 5]); await capture(page, 'jsdoc-pagination-known');
  panel = await scenario('Pagination', pilotTitles.Pagination[1]);
  await panel.getByRole('button', { name: 'Next page', exact: true }).click();
  check('Pending retains actual rows', await rows(panel), [1, 2, 3, 4, 5]);
  check('Pending blocks duplicates', await panel.getByRole('button', { name: 'Next page', exact: true }).isDisabled(), true);
  await panel.getByRole('button', { name: 'Complete pending request', exact: true }).click();
  check('Completion applies actual rows', await rows(panel), [6, 7, 8, 9, 10]);
  await panel.getByRole('button', { name: 'Next page', exact: true }).click(); await restore(panel, 'Pending rows');
  check('Reset invalidates pending completion', await panel.getByRole('button', { name: 'Complete pending request', exact: true }).isDisabled(), true);
  check('Pending reset rows', await rows(panel), [1, 2, 3, 4, 5]); await capture(page, 'jsdoc-pagination-pending');
  panel = await scenario('Pagination', pilotTitles.Pagination[2]);
  check('Empty rows', await rows(panel), []);
  check('Empty next disabled', await panel.getByRole('button', { name: 'Next page', exact: true }).isDisabled(), true);
  await restore(panel, 'Empty rows'); check('Empty stays empty', await rows(panel), []); await capture(page, 'jsdoc-pagination-empty');

  panel = await scenario('Pagination', pilotTitles.Pagination[3]);
  const advancedRows = () => panel.locator('[data-order-id]').evaluateAll(nodes => nodes.map(node => node.getAttribute('data-order-id')));
  for (const mode of ['known','unknown','cursor','loadMore']) {
    await panel.locator('select[data-pagination-demo=mode]').selectOption(mode);
    check(`Advanced ${mode}: baseline rows`, await advancedRows(), ['order-1','order-2','order-3','order-4','order-5']);
    await panel.getByRole('button', { name: mode === 'loadMore' ? 'Load additional orders' : 'Next page', exact: true }).click();
    check(`Advanced ${mode}: actual rows change or append`, await advancedRows(), mode === 'loadMore'
      ? ['order-1','order-2','order-3','order-4','order-5','order-6','order-7','order-8','order-9','order-10']
      : ['order-6','order-7','order-8','order-9','order-10']);
  }
  await panel.getByRole('checkbox', { name: 'Complete requests manually', exact: true }).check();
  await panel.getByRole('button', { name: 'Load additional orders', exact: true }).click();
  check('Advanced pending locks native nav', await panel.locator('nav').getAttribute('aria-busy'), 'true');
  check('Advanced pending retains ten real orders', (await advancedRows()).length, 10);
  await panel.getByRole('button', { name: 'Complete loading', exact: true }).click();
  check('Advanced completion appends exactly one batch', (await advancedRows()).length, 15);
  await panel.getByRole('button', { name: 'Load additional orders', exact: true }).click();
  await panel.locator('select[data-pagination-demo=filter]').selectOption('missing');
  check('Filter clears actual pending state', await panel.getByRole('button', { name: 'Complete loading', exact: true }).isDisabled(), true);
  check('Advanced empty result has no rows', await advancedRows(), []);
  await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click();
  await panel.getByRole('checkbox', { name: 'Hide navigation', exact: true }).check();
  await panel.getByRole('button', { name: 'Inspect navigation ref', exact: true }).click();
  await text(panel, 'Native ref: NAV; hidden=true; busy=absent');
  await panel.getByRole('checkbox', { name: 'Hide navigation', exact: true }).uncheck();
  await panel.getByRole('checkbox', { name: 'RTL override (unchecked inherits host)', exact: true }).check();
  check('Advanced explicit RTL reaches actual nav', await panel.locator('nav').evaluate(node => getComputedStyle(node).direction), 'rtl');
  await restore(panel, 'Advanced modes');
  check('Advanced reset restores five orders', (await advancedRows()).length, 5);
  check('Advanced reset releases RTL', await panel.getByRole('checkbox', { name: 'RTL override (unchecked inherits host)', exact: true }).isChecked(), false);
  await capture(page, 'jsdoc-pagination-advanced');

  panel = await scenario('useSx', pilotTitles.useSx[0]);
  const adapterClasses = [];
  for (const adapter of ['root','domain','legacy']) {
    await panel.locator('select[data-demo-variant=imports]').selectOption(adapter);
    check(`${adapter}: one actual adapter`, await panel.locator('[data-imports-demo=root],[data-imports-demo=domain],[data-imports-demo=legacy]').count(), 1);
    await panel.getByRole('button', { name: 'Increment counter', exact: true }).click();
    await panel.getByRole('button', { name: 'Invoke captured callback', exact: true }).click();
    await text(panel, 'Observed counter: 1'); await text(panel, 'Same callback identity: yes');
    const preview = panel.locator(`[data-imports-demo="${adapter}-preview"]`);
    adapterClasses.push(await preview.getAttribute('class'));
    check(`${adapter}: real baseline width`, await preview.evaluate(node => getComputedStyle(node).width), '240px');
    await panel.getByRole('button', { name: 'Apply alternate width', exact: true }).click();
    check(`${adapter}: real composed width`, await preview.evaluate(node => getComputedStyle(node).width), '360px');
    await panel.getByRole('button', { name: 'Remove alternate width', exact: true }).click();
    check(`${adapter}: width override removed`, await preview.evaluate(node => getComputedStyle(node).width), '240px');
    await restore(panel, `${adapter} imports`);
    await panel.getByRole('button', { name: 'Invoke captured callback', exact: true }).click();
    await text(panel, 'Observed counter: 0'); await text(panel, 'Same callback identity: yes');
  }
  check('Import adapters share actual composed classes', new Set(adapterClasses).size, 1);
  await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click();
  await panel.getByRole('button', { name: 'Override direction to RTL', exact: true }).click();
  check('Imports explicit direction reaches real provider', await panel.locator('[data-imports-demo=provider]').evaluate(node => getComputedStyle(node).direction), 'rtl');
  await panel.getByRole('button', { name: 'Use host direction', exact: true }).click();
  await page.setViewportSize({ width: 320, height: 900 }); await settle(page);
  await panel.getByRole('button', { name: 'Apply alternate width', exact: true }).click();
  const scroll = panel.locator('[data-imports-demo=legacy-scroll]');
  check('Imports real fixed width scrolls in bounded region', await scroll.evaluate(node => node.scrollWidth > node.clientWidth && getComputedStyle(node).overflowX === 'auto' && node.tabIndex === 0), true);
  await overflow(page, 'Import interoperability 320px'); await capture(page, 'jsdoc-imports-320');
  await page.setViewportSize({ width: 1440, height: 1100 }); await settle(page);

  const nativeCalls = () => page.evaluate(() => window.trimFixture.calls.length);
  for (const [name, titles] of Object.entries(pilotTitles).filter(([name]) => name.startsWith('useSPFx') && name !== 'useSPFxPnPList')) {
    for (const title of titles) {
      const before = await nativeCalls(); panel = await scenario(name, title);
      check(`${title}: initial selection performs no native request`, await nativeCalls(), before);
      check(`${title}: finite configuration result`, await panel.locator('[data-demo-result]').count(), 1);
      await restore(panel, title); check(`${title}: reset performs no native request`, await nativeCalls(), before);
    }
  }
  panel = await scenario('useSPFxLocaleInfo', pilotTitles.useSPFxLocaleInfo[0]);
  await panel.getByRole('combobox', { name: 'Period', exact: true }).selectOption('summer');
  check('Context summer is one actual time', await panel.locator('time').count(), 1);
  check('Context summer exact instant', await panel.locator('time').getAttribute('datetime'), '2026-07-15T12:00:00.000Z');
  await restore(panel, 'Context seasonal');
  check('Context reset restores winter', await panel.locator('time').getAttribute('datetime'), '2026-01-15T12:00:00.000Z');
  panel = await scenario('useSPFxCrossSitePermissions', pilotTitles.useSPFxCrossSitePermissions[0]);
  const legacyReads = await nativeCalls(); await panel.getByRole('textbox', { name: 'Site URL', exact: true }).fill('https://simulated.invalid/sites/target');
  check('Context legacy draft performs no read', await nativeCalls(), legacyReads);
  await panel.getByRole('button', { name: 'Check / repeat', exact: true }).click();
  await checkLegacyPermissionResult(panel, check);
  check('Context legacy checks actual site and web masks', (await nativeCalls()) >= legacyReads + 2, true);
  panel = await scenario('useSPFxHubSiteInfo', pilotTitles.useSPFxHubSiteInfo[0]);
  const hubReads = await nativeCalls(); await panel.getByRole('button', { name: 'Start / repeat lookup', exact: true }).click();
  await panel.getByText('Hub URL: https://simulated.invalid/sites/fixture-hub', { exact: true }).waitFor();
  check('Context Hub actual native read after Start', await nativeCalls(), hubReads + 1);
  await restore(panel, 'Context Hub');
  panel = await scenario('useStableCallback', pilotTitles.useStableCallback[0]);
  await panel.getByRole('button', { name: 'Increment counter', exact: true }).click(); await panel.getByRole('button', { name: 'Increment counter', exact: true }).click();
  await panel.getByRole('button', { name: 'Invoke original callback', exact: true }).click();
  await text(panel, 'Observed counter: 2'); await text(panel, 'Same callback identity: yes');
  await restore(panel, 'Current callback'); await panel.getByRole('button', { name: 'Invoke original callback', exact: true }).click();
  await text(panel, 'Observed counter: 0'); await text(panel, 'Same callback identity: yes'); await capture(page, 'jsdoc-callback-current');
  panel = await scenario('useStableCallback', pilotTitles.useStableCallback[1]);
  await panel.getByRole('button', { name: 'Rerender component', exact: true }).click(); await panel.getByRole('button', { name: 'Rerender component', exact: true }).click();
  await panel.getByRole('button', { name: 'Invoke original callback', exact: true }).click(); await text(panel, 'Rerender revision: 2'); await text(panel, 'Same callback identity: yes');
  await restore(panel, 'Callback identity'); await panel.getByRole('button', { name: 'Invoke original callback', exact: true }).click(); await text(panel, 'Same callback identity: yes'); await capture(page, 'jsdoc-callback-identity');

  const before = await calls(); panel = await scenario('useSPFxPnPList', pilotTitles.useSPFxPnPList[0]);
  check('Read: selecting feature issues zero requests', await calls(), before);
  const listTitle = panel.getByRole('textbox', { name: 'List title', exact: true });
  await listTitle.fill('Documents'); check('Read: resource editing issues zero requests', await calls(), before);
  await panel.getByRole('button', { name: 'Read selected list', exact: true }).click();
  await panel.getByText('Results for Documents: 1', { exact: true }).waitFor();
  check('Explicit read performs one request', await calls(), before + 1);
  await page.evaluate(() => window.pnpReadFixture.configure({ failure: 'unavailable', hold: false }));
  await panel.getByRole('button', { name: 'Read selected list', exact: true }).click(); await panel.getByRole('alert').filter({ hasText: 'Request error for Documents' }).waitFor();
  await page.evaluate(() => window.pnpReadFixture.configure({ failure: undefined, hold: true }));
  await panel.getByRole('button', { name: 'Read selected list', exact: true }).click();
  await panel.getByText('Loading Documents…', { exact: true }).waitFor(); await restore(panel, 'Read request');
  await page.evaluate(() => window.pnpReadFixture.release()); await settle(page);
  check('Reset clears selected list', await listTitle.inputValue(), '');
  check('Late read cannot republish result', await panel.getByRole('list').count(), 0);
  check('Fixture forbids every write', await page.evaluate(() => window.pnpReadFixture.calls.every(call => call.method === 'GET')), true);
  await capture(page, 'jsdoc-list-read');

  await page.setViewportSize({ width: 320, height: 900 }); await settle(page);
  const mobile = page.getByRole('combobox', { name: 'Funzionalità', exact: true });
  check('320px mobile feature selector visible', await mobile.isVisible(), true);
  check('320px desktop navigation hidden', await navigation.isVisible(), false);
  check('Mobile optgroups', await mobile.locator('optgroup').evaluateAll(nodes => nodes.map(node => node.label)), ['Controlli', 'Hook']);
  const pagination = await mobile.locator('option').filter({ hasText: /^Pagination$/ }).getAttribute('value');
  await mobile.selectOption(pagination); await mobile.focus(); await page.keyboard.press('Tab');
  check('Keyboard reaches central scenario selector', await page.getByRole('combobox', { name: 'Scenario', exact: true }).evaluate(element => element === document.activeElement), true);
  await overflow(page, 'JSDoc 320px'); await capture(page, 'jsdoc-mobile-320');
}
module.exports = { jsdocPilotsBrowser, checkLegacyPermissionResult };
