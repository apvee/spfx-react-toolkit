// Browser plugin unavailable: installed Playwright + Chrome; actual demo with mocked hook.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const { teamsDarkTheme } = require('@fluentui/react-components');
const root = path.resolve(__dirname, '..');
const directory = path.join(root, 'temp/pnp-confirmation-browser');
const output = path.resolve(process.env.PNP_CONFIRMATION_EVIDENCE || '.docs/maintenance/evidence/pnp-confirmation');
const checks = [], diagnostics = [], network = [], screenshots = [];
const check = (name, actual, expected) => { checks.push({ name, actual, expected }); assert.deepEqual(actual, expected, name); };
const server = http.createServer((request, response) => {
  const pathname = new URL(request.url, 'http://127.0.0.1').pathname;
  if (pathname === '/favicon.ico') return response.writeHead(204).end();
  const file = path.join(directory, pathname === '/' ? 'index.html' : pathname.slice(1));
  if (!file.startsWith(`${directory}/`) || !fs.existsSync(file)) return response.writeHead(404).end();
  response.setHeader('Content-Type', file.endsWith('.html') ? 'text/html' : 'application/javascript'); response.end(fs.readFileSync(file));
});
(async () => {
  fs.mkdirSync(output, { recursive: true }); await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  let browser;
  try {
    browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
    const page = await browser.newPage({ viewport: { width: 640, height: 850 } });
    page.on('pageerror', error => diagnostics.push(error.message));
    page.on('console', message => { if (['warning', 'error'].includes(message.type())) diagnostics.push(message.text()); });
    page.on('dialog', dialog => { diagnostics.push(`Unexpected native ${dialog.type()}`); dialog.dismiss(); });
    const url = `http://127.0.0.1:${server.address().port}`;
    await page.route('**/*', route => { if (new URL(route.request().url()).origin === url) return route.continue(); network.push(route.request().url()); return route.abort(); });
    await page.goto(url); check('Page identity', await page.title(), 'Mocked PnP confirmation checkpoint');
    check('Meaningful actual demo', await page.getByRole('heading', { name: 'Example 3: useSPFxPnPList - CRUD Operations' }).count(), 1);
    check('One selected service result', await page.locator('[data-demo-section] [data-demo-result]').count(), 1);
    check('Framework overlays absent', await page.locator('webpack-dev-server-client-overlay,#webpack-dev-server-client-overlay').count(), 0);
    const trigger = page.getByRole('button', { name: 'Delete', exact: true }).first();
    const dialog = page.getByRole('dialog', { name: 'Delete this item?' });
    await page.getByLabel('List Title', { exact: true }).fill('Mock list');
    for (const action of ['Cancel', 'Escape']) {
      await trigger.click(); await dialog.waitFor(); check(`${action}: intended item`, (await dialog.textContent()).includes('Delete item #42?'), true);
      check(`${action}: consent before removal`, await page.evaluate(() => window.pnpConfirmationMock.calls.length), 0);
      if (action === 'Cancel') await dialog.getByRole('button', { name: 'Cancel', exact: true }).click(); else await page.keyboard.press('Escape');
      await dialog.waitFor({ state: 'hidden' });
      check(`${action}: no removal`, await page.evaluate(() => window.pnpConfirmationMock.calls.length), 0);
      check(`${action}: origin focus restored`, await trigger.evaluate(element => element === document.activeElement), true);
    }
    await page.evaluate(() => window.demoFixtureHost.update({ theme: 'dark', dir: 'rtl' }));
    await trigger.click(); await dialog.waitFor();
    await page.waitForFunction(() => { const element = document.querySelector('[role=dialog]'); return element && getComputedStyle(element).opacity === '1'; });
    check('Confirmation inherits dark token', await dialog.evaluate(element => getComputedStyle(element).getPropertyValue('--colorNeutralBackground1').trim()), teamsDarkTheme.colorNeutralBackground1);
    check('Confirmation inherits RTL', await dialog.evaluate(element => getComputedStyle(element).direction), 'rtl');
    check('Modal owns keyboard focus', await dialog.evaluate(element => element.contains(document.activeElement)), true);
    const screenshot = path.join(output, 'confirmation-dark-rtl.png'); await page.screenshot({ path: screenshot }); screenshots.push(screenshot);
    check('Scenario locked during consent', await page.locator('select[data-demo-variant=PnPPanel]').isDisabled(), true);
    check('Target read-only during consent', await page.getByLabel('List Title', { exact: true }).getAttribute('readonly') !== null, true);
    await page.evaluate(() => window.pnpConfirmationMock.replaceOperation());
    await dialog.getByRole('button', { name: 'Delete item', exact: true }).click(); await dialog.waitFor({ state: 'hidden' });
    check('Captured ID and operation dispatch exactly once', await page.evaluate(() => window.pnpConfirmationMock.calls), [{ id: 42, listTitle: 'Mock list', operation: 0 }]);
    check('Scenario locked during pending removal', await page.locator('select[data-demo-variant=PnPPanel]').isDisabled(), true);
    check('Target read-only during pending removal', await page.getByLabel('List Title', { exact: true }).getAttribute('readonly') !== null, true);
    await page.evaluate(() => window.pnpConfirmationMock.settle());
    check('Scenario unlocked after mutation settles', await page.locator('select[data-demo-variant=PnPPanel]').isDisabled(), false);
    check('Mutation resolves before refetch removes origin', await page.getByText('#42 - Intended item', { exact: true }).count(), 1);
    check('Approval uses stable list focus before refetch', await page.getByLabel('List Title', { exact: true }).evaluate(element => document.activeElement === element), true);
    await page.evaluate(() => window.pnpConfirmationMock.publishRows());
    await page.getByText('#42 - Intended item', { exact: true }).waitFor({ state: 'hidden' });
    await page.waitForFunction(() => document.activeElement?.getAttribute('placeholder') === 'e.g., Site Pages');
    check('Deleted origin restores useful list field focus', await page.getByLabel('List Title', { exact: true }).evaluate(element => document.activeElement === element), true);
    check('Other intended item remains', await page.getByText('#73 - Other item', { exact: true }).count(), 1);
    await page.getByRole('button', { name: 'Delete', exact: true }).click(); await dialog.waitFor();
    await dialog.getByRole('button', { name: 'Delete item', exact: true }).click(); await dialog.waitFor({ state: 'hidden' });
    const destination = page.getByLabel('New item title...', { exact: true }); await destination.click();
    check('User chooses another focus destination', await destination.evaluate(element => document.activeElement === element), true);
    await page.evaluate(() => window.pnpConfirmationMock.settle());
    check('Later mutation completion preserves chosen focus', await destination.evaluate(element => document.activeElement === element), true);
    check('Second mutation also resolves before refetch', await page.getByText('#73 - Other item', { exact: true }).count(), 1);
    await page.evaluate(() => window.pnpConfirmationMock.publishRows()); await page.getByText('#73 - Other item', { exact: true }).waitFor({ state: 'hidden' });
    check('Later row publication preserves chosen focus', await destination.evaluate(element => document.activeElement === element), true);
    check('Second explicit operation dispatches once', await page.evaluate(() => window.pnpConfirmationMock.calls), [{ id: 42, listTitle: 'Mock list', operation: 0 }, { id: 73, listTitle: 'Mock list', operation: 1 }]);
    check('All actual network remains loopback', network, []); check('No unexpected browser diagnostics/native dialogs', diagnostics, []);
    console.log(`PnP confirmation: ${checks.length} PASS checks; actual Fluent surface and mocked operation only.`);
  } finally {
    fs.writeFileSync(path.join(output, 'browser-results.json'), JSON.stringify({ checks, diagnostics, network, screenshots, boundary: 'actual demo/Fluent Dialog; mocked PnP hook/removal; no actual CRUD, network or authenticated host', browserPlugin: 'not available' }, undefined, 2));
    await browser?.close(); await new Promise(resolve => server.close(resolve));
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
