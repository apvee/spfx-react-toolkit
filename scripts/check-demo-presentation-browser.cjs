// Browser plugin unavailable: existing Playwright + installed Chrome. Actual app/lazy registry.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const { teamsDarkTheme, webLightTheme } = require('@fluentui/react-components');
const { groupsProduct } = require('../tests/fixtures/demo-control-isolation/task3-product.browser.cjs');
const root = path.resolve(__dirname, '..');
const directory = path.join(root, process.env.DEMO_PRESENTATION_MODE === 'source' ? 'temp/demo-presentation-source-browser' : 'temp/demo-presentation-browser');
const output = path.resolve(process.env.DEMO_PRESENTATION_EVIDENCE || path.join(root, '.docs/maintenance/evidence/demo-presentation'));
const executablePath = process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const registry = fs.readFileSync(path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demoRegistry.ts'), 'utf8');
const entries = [...registry.matchAll(/key: '([^']+)',\s*title: '([^']+)'/g)].map(match => ({ key: match[1], title: match[2] }));
const checks = [], diagnostics = [], metrics = [], screenshots = [], fontAssets = [];
const check = (name, actual, expected) => { checks.push({ name, actual, expected }); assert.deepEqual(actual, expected, name); };
const server = http.createServer((request, response) => {
  const pathname = new URL(request.url, 'http://127.0.0.1').pathname;
  if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
  if (pathname === '/cards-demo-missing-image.png') { response.writeHead(200, { 'Content-Type': 'image/png' }).end('invalid image'); return; }
  const file = path.join(directory, pathname === '/' ? 'index.html' : pathname.slice(1));
  if (!file.startsWith(`${directory}/`) || !fs.existsSync(file)) { response.writeHead(404).end(); return; }
  response.setHeader('Content-Type', file.endsWith('.html') ? 'text/html' : 'application/javascript'); response.end(fs.readFileSync(file));
});
const observe = page => {
  page.on('pageerror', error => diagnostics.push({ type: 'pageerror', text: error.message }));
  page.on('console', message => { if (['error', 'warning'].includes(message.type())) diagnostics.push({ type: message.type(), text: message.text() }); });
};
const settle = page => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
async function guardNetwork(page, url) {
  await page.route('**/*', async route => { const request = new URL(route.request().url()); if (/^https:\/\/res-1\.cdn\.office\.net\/files\/fabric-cdn-prod_[^/]+\/assets\/fonts\/segoeui-westeuropean\/segoeui-(regular|semibold|bold)\.woff2?$/.test(request.href)) { fontAssets.push({ requested: request.href, localSource: '/System/Library/Fonts/Supplemental/Arial.ttf', boundary: 'Unavailable external legacy text font; local system substitute' }); return route.fulfill({ contentType: 'font/ttf', body: fs.readFileSync('/System/Library/Fonts/Supplemental/Arial.ttf') }); } if (request.origin !== url && !['data:', 'blob:'].includes(request.protocol)) { diagnostics.push({ type: 'network', text: request.href }); return route.abort(); } return route.continue(); });
}
async function select(page, key) {
  const entry = entries.find(item => item.key === key);
  const tab = page.getByRole('tab', { name: entry.title, exact: true });
  await tab.click();
  await page.waitForFunction(() => document.querySelector('[role=tabpanel] h3, [role=tabpanel] [data-shell-demo=panel]'));
  await settle(page);
  check(`${key}: selected`, await tab.getAttribute('aria-selected'), 'true');
  check(`${key}: tabpanel linkage`, await page.getByRole('tabpanel').getAttribute('aria-labelledby'), await tab.getAttribute('id'));
}
async function capture(page, name) {
  const file = path.join(output, `${name}.png`); await page.screenshot({ path: file, fullPage: true }); screenshots.push(file);
}
async function overflow(page, name) {
  const value = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, client: document.documentElement.clientWidth }));
  if (value.scroll > value.client + 1) fs.writeFileSync(path.join(output, 'overflow-debug.json'), JSON.stringify(await page.evaluate(() => [...document.querySelectorAll('body *')].filter(e => e.getClientRects().length && e.getBoundingClientRect().right > innerWidth + 1).map(e => ({ tag: e.tagName, classes: e.className, text: e.textContent?.slice(0,100), width: e.getBoundingClientRect().width, left: e.getBoundingClientRect().left })).slice(-30)), undefined, 2));
  check(`${name}: page overflow`, value.scroll <= value.client + 1, true);
}
async function labels(page, name) {
  const missing = await page.getByRole('tabpanel').evaluate(panel => [...panel.querySelectorAll('input:not([type=hidden]),select,textarea,button')].filter(element => {
    if (!element.getClientRects().length || element.closest('[inert]') || getComputedStyle(element).visibility === 'hidden') return false;
    const labelled = element.getAttribute('aria-labelledby');
    const name = element.getAttribute('aria-label') || (labelled && labelled.split(' ').map(id => document.getElementById(id)?.textContent || '').join(' ')) || [...element.labels || []].map(label => label.textContent).join(' ') || (element.tagName === 'BUTTON' ? element.textContent : '');
    return !name.trim();
  }).map(element => `${element.tagName}:${element.getAttribute('name') || element.id || element.outerHTML.slice(0, 150)}`));
  check(`${name}: visible native controls have accessible labels`, missing, []);
}
async function singleExample(page, key, checkpoint = 'initial') {
  const name = `${key}/${checkpoint}`;
  const panel = page.getByRole('tabpanel');
  check(`${name}: one demo section`, await panel.locator('[data-demo-section]').count(), 1);
  check(`${name}: one configuration card`, await panel.locator('[data-demo-configuration]').count(), 1);
  check(`${name}: one principal subject or current result`, await panel.locator('[data-demo-subject],[data-demo-result]').count(), 1);
  const principalRoots = {
    cards: '.fui-Card:not([data-demo-configuration])',
    form: 'form',
    'data-grid': '.fui-DataGrid',
    'description-list': 'dl',
    groups: '[data-groups-demo=settings],[data-groups-demo=accordion]',
  };
  if (principalRoots[key]) check(`${name}: one actual principal control, including retained diagnostics`, await panel.locator(principalRoots[key]).count(), 1);
}
// Select only explicit principal variants: ordinary property controls are covered by
// panel behavior checks, rather than an exhaustive combination of every setting.
async function singleExampleTransitions(page, key, inspectVariant = async () => {}) {
  const panel = page.getByRole('tabpanel');
  const diagnostics = panel.getByRole('button', { name: 'Advanced diagnostics', exact: true });
  const openDiagnostics = async checkpoint => {
    if (await diagnostics.count()) {
      if (await diagnostics.getAttribute('aria-expanded') !== 'true') await diagnostics.click();
      await settle(page);
      check(`${key}/${checkpoint}: diagnostics open`, await diagnostics.getAttribute('aria-expanded'), 'true');
      await singleExample(page, key, checkpoint);
    }
  };
  await singleExample(page, key);
  await openDiagnostics('diagnostics');
  // Re-read selectors in each branch: selecting a scenario can introduce a
  // second operation selector that was absent in the default branch.
  const walk = async (visited, chain) => {
    const names = await panel.locator('select[data-demo-variant]').evaluateAll(nodes => nodes.filter(node => !node.disabled).map(node => node.getAttribute('data-demo-variant')));
    const name = names.find(candidate => !visited.has(candidate));
    if (name === undefined) return;
    const variant = panel.locator(`select[data-demo-variant=${JSON.stringify(name)}]`);
    const original = await variant.inputValue();
    const options = await variant.locator('option').evaluateAll(nodes => nodes.filter(option => !option.disabled && !option.parentElement.disabled).map(option => option.value));
    for (const value of options) {
      await variant.selectOption(value); await settle(page);
      const checkpoint = [...chain, `${name}=${value}`].join('/');
      await singleExample(page, key, checkpoint);
      await inspectVariant(checkpoint);
      await openDiagnostics(`${checkpoint}/diagnostics`);
      await walk(new Set([...visited, name]), [...chain, `${name}=${value}`]);
    }
    await variant.selectOption(original); await settle(page);
    await singleExample(page, key, `${name}/restored`);
  };
  await walk(new Set(), []);
  if (await diagnostics.count() && await diagnostics.getAttribute('aria-expanded') === 'true') {
    await diagnostics.click();
    await settle(page);
    await singleExample(page, key, 'diagnostics-closed');
  }
}
async function interact(page, key) {
  const panel = page.getByRole('tabpanel');
  switch (key) {
    case 'cards': await panel.locator('[data-card-instance=main] [data-cards-demo=title-input]').fill('Integrated title'); check('Cards editable title', await panel.locator('[data-card-instance=main] [data-cards-demo=preview-card] h2').textContent(), 'Integrated title'); break;
    case 'web-part-shell': await panel.locator('[data-shell-demo=note]').fill('Retained shell note'); await panel.locator('[data-shell-demo=increment]').click(); check('Shell counter', (await panel.locator('[data-shell-demo=observations]').textContent()).includes('"count":1'), true); break;
    case 'description-list': { const select = panel.locator('[data-description-list-demo=layout]').first(); await select.selectOption('vertical'); check('DescriptionList actual vertical term/value layout', await panel.locator('[data-description-list-demo=list]').first().evaluate(list => { const term = list.querySelector('dt').getBoundingClientRect(), value = list.querySelector('dd').getBoundingClientRect(); return value.top >= term.bottom - 1; }), true); break; }
    case 'pagination': { const before = await panel.locator('[data-pagination-demo=summary]').first().textContent(); await panel.getByRole('button', { name: 'Next page', exact: true }).first().click(); check('Pagination actual results change', await panel.locator('[data-pagination-demo=summary]').first().textContent() !== before, true); break; }
    case 'data-grid': { const row = panel.locator('#data-grid-main .fui-DataGridBody > [role=row]').first(); const before = (await row.boundingBox()).height; await panel.locator('#dg-size').selectOption('small'); check('DataGrid actual row height follows size', (await row.boundingBox()).height < before, true); break; }
    case 'form': {
      const title = panel.getByLabel('Title', { exact: false });
      await title.fill(''); await panel.locator('[data-form-demo=submit]').click();
      check('Form actual required rule blocks submit', (await panel.locator('[data-form-demo=status]').textContent()).includes('Controlla'), true);
      await title.fill('Report integrato'); await panel.locator('[data-form-demo=submit]').click();
      await page.waitForFunction(() => document.querySelector('[data-form-demo=status]')?.textContent.includes('Saved'));
      check('Form actual valid record saved', (await panel.locator('[data-form-demo=status]').textContent()).includes('Saved'), true); break;
    }
    case 'stage': await panel.locator('[data-stage-demo=mode]').first().selectOption('popover'); await panel.locator('[data-stage-demo=trigger]').first().click(); await page.locator('.fui-PopoverSurface').waitFor(); check('Stage actual popover presentation', await page.locator('.fui-PopoverSurface').count(), 1); await page.getByRole('button', { name: 'Close A', exact: true }).click(); await page.locator('.fui-PopoverSurface').waitFor({ state: 'hidden' }); break;
    case 'groups': await groupsProduct(page, { check }); break;
    case 'state-message': { const artwork = panel.locator('[data-state-message-demo=preview] svg'); const before = await artwork.innerHTML(); await panel.locator('[data-state-message-demo=preset]').selectOption('error'); check('StateMessage actual preset artwork changes', await artwork.innerHTML() !== before, true); await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click(); await panel.locator('[data-state-message-demo=role]').check(); check('StateMessage explicit alert applies to text region', await panel.locator('[data-state-message-demo=preview] [data-state-message-text]').getAttribute('role'), 'alert'); break; }
    case 'auto-skeleton': { const note = panel.locator('[data-auto-skeleton-demo=focus-target]'); await note.fill('Retained loading note'); await panel.locator('[data-auto-skeleton-demo=loading]').check(); check('AutoSkeleton actual busy state', await panel.locator('[data-auto-skeleton-demo=preview]').getAttribute('aria-busy'), 'true'); await panel.locator('[data-auto-skeleton-demo=loading]').uncheck(); check('AutoSkeleton retained note', await note.inputValue(), 'Retained loading note'); break; }
    case 'relative-time': { const locale = panel.locator('[data-relative-time-demo=locale]'); const preview = panel.locator('[data-relative-time-demo=preview]'); const before = await preview.textContent(); await locale.selectOption('it-IT'); check('RelativeTime actual translated text', await preview.textContent() !== before, true); await locale.selectOption('en-US'); break; }
    case 'formatted-text': { const field = panel.locator('[data-formatted-text-demo=text]'); await field.fill('Integrated **markdown**'); check('FormattedText actual parsed content', await panel.locator('[data-formatted-text-demo=preview] strong').textContent(), 'markdown'); break; }
    case 'imports': {
      await panel.getByRole('button', { name: 'Increment counter', exact: true }).click();
      await panel.getByRole('button', { name: 'Invoke captured callback', exact: true }).click();
      const diagnostics = panel.getByRole('button', { name: 'Advanced diagnostics', exact: true });
      await diagnostics.click(); check('Imports original callback observes latest state', (await panel.textContent()).includes('Observed counter: 1'), true); await diagnostics.click(); break;
    }
    case 'react-hooks': await panel.getByRole('button', { name: 'Increment counter', exact: true }).click(); await panel.getByRole('button', { name: 'Invoke captured callback', exact: true }).click(); check('Stable callback identity', (await panel.textContent()).includes('Same callback identity: yes'), true); break;
    case 'styles': {
      const scenario = panel.locator('select[data-demo-variant=styles]');
      await scenario.selectOption('preset');
      const preview = panel.locator('[data-sx-demo=preset-preview]'); const before = await preview.evaluate(element => getComputedStyle(element).backgroundColor);
      await panel.locator('[data-sx-demo=preset]').selectOption('alternative'); check('Styles actual preset background changes', await preview.evaluate(element => getComputedStyle(element).backgroundColor) !== before, true);
      for (const [mode, input, name] of [['layout', 'layout-width', 'Layout preview'], ['catalog', 'catalog-value', 'Catalog preview'], ['container', 'query-a-width', 'Container query previews']]) {
        await scenario.selectOption(mode);
        if (mode === 'catalog') { await panel.locator('[data-sx-demo=catalog-family]').selectOption('minWidth'); await panel.locator('[data-sx-demo=catalog-member]').selectOption('px'); }
        await panel.locator(`[data-sx-demo=${input}]`).fill('2000');
        const region = panel.getByRole('region', { name, exact: true });
        check(`Styles ${mode}: focusable`, await region.getAttribute('tabindex'), '0');
        check(`Styles ${mode}: actual overflow bounded`, await region.evaluate(element => element.scrollWidth > element.clientWidth && getComputedStyle(element).overflowX === 'auto'), true);
        await region.focus(); const initial = await region.evaluate(element => element.scrollLeft); await page.keyboard.press('ArrowRight');
        await page.waitForFunction(({ name, initial }) => [...document.querySelectorAll('[role=region]')].find(element => element.getAttribute('aria-label') === name)?.scrollLeft > initial, { name, initial });
      }
      break;
    }
    case 'providers': await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click(); check('Provider disclosure', await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).getAttribute('aria-expanded'), 'true'); break;
    case 'context': await panel.locator('select[data-demo-variant=context]').selectOption('cross-site'); await panel.locator('[data-context-demo=url-draft]').fill('https://simulated.invalid/sites/target'); check('Context staged URL draft', await panel.locator('[data-context-demo=url-draft]').inputValue(), 'https://simulated.invalid/sites/target'); break;
    case 'runtime': await panel.getByLabel('Description', { exact: true }).fill('Integrated properties'); await panel.getByRole('button', { name: 'setProperties', exact: true }).click(); check('Runtime actual provider properties', (await panel.textContent()).includes('Current Description:Integrated properties'), true); break;
    case 'clients': await panel.locator('select[data-demo-variant=ClientsPanel]').selectOption('aad'); await panel.getByLabel('Resource URL or App ID URI', { exact: true }).fill('api://simulated'); check('Clients configuration draft', await panel.getByLabel('Resource URL or App ID URI', { exact: true }).inputValue(), 'api://simulated'); break;
    case 'graph': await panel.getByRole('button', { name: 'Run', exact: true }).click(); await panel.getByText('Simulated network unavailable', { exact: true }).waitFor(); check('Graph simulated read error visible', await panel.getByText('Simulated network unavailable', { exact: true }).count(), 1); break;
    case 'site': case 'tenant': await panel.locator(`select[data-demo-variant=${key === 'site' ? 'SitePanel' : 'TenantPanel'}]`).selectOption(key === 'site' ? 'save' : 'store'); if (key === 'tenant') await panel.locator('select[data-demo-variant=TenantStoreScenario]').selectOption('save'); await panel.getByLabel('Value', { exact: true }).fill('Local draft'); check(`${key}: editable draft`, await panel.getByLabel('Value', { exact: true }).inputValue(), 'Local draft'); break;
    case 'pnp': await panel.locator('select[data-demo-variant=PnPPanel]').selectOption('listAccess'); await panel.getByLabel('List GUID', { exact: true }).fill('11111111-2222-3333-4444-555555555555'); check('PnP GUID config', await panel.getByLabel('List GUID', { exact: true }).inputValue(), '11111111-2222-3333-4444-555555555555'); break;
    case 'search': await panel.locator('select[data-demo-variant=SearchPanel]').selectOption('suggestions'); await panel.getByLabel('Search with Autocomplete', { exact: true }).fill('local'); await panel.getByRole('button', { name: 'local simulated suggestion', exact: true }).waitFor(); check('Search real debounce/suggestion UI', await panel.getByRole('button', { name: 'local simulated suggestion', exact: true }).count(), 1); break;
    default: throw new Error(`Uncovered panel ${key}`);
  }
}
async function liveTheme(page) {
  await select(page, 'cards');
  const main = page.locator('[data-card-instance=main]'), note = main.locator('[data-cards-demo=retained-input]');
  await main.locator('[data-cards-demo=header-action]').click();
  await note.fill('Live inherited theme draft'); const identity = await note.elementHandle(); await note.focus();
  for (const patch of [{ theme: 'dark' }, { theme: 'custom', brandColor: '#7b2cbf' }, { theme: 'custom', brandColor: '#cf2468', dir: 'rtl' }, { theme: 'light', dir: 'ltr' }]) {
    await page.evaluate(value => window.demoFixtureHost.update(value), patch); await settle(page);
    check(`Live theme ${JSON.stringify(patch)} retains node/value/focus`, await identity.evaluate(element => element === document.querySelector('[data-card-instance=main] [data-cards-demo=retained-input]') && element.value === 'Live inherited theme draft' && document.activeElement === element), true);
    check(`Live theme ${JSON.stringify(patch)} retains counter`, (await main.locator('[data-cards-demo=counters]').textContent()).includes('Header:1'), true);
    if (patch.theme === 'dark' || patch.theme === 'light') {
      const theme = patch.theme === 'dark' ? teamsDarkTheme : webLightTheme;
      for (const name of ['colorNeutralBackground1', 'colorNeutralForeground1', 'colorBrandBackground'])
        check(`Inherited ${patch.theme} ${name}`, await main.evaluate((element, name) => getComputedStyle(element).getPropertyValue(`--${name}`).trim(), name), theme[name]);
    }
    if (patch.brandColor) check(`Custom brand token ${patch.brandColor}`, await main.evaluate(element => getComputedStyle(element).getPropertyValue('--colorBrandBackground').trim()), patch.brandColor);
    if (patch.dir) check(`Host direction ${patch.dir} inherited`, await main.evaluate(element => getComputedStyle(element).direction), patch.dir);
  }
  await identity.dispose();
  check('Visible app has no duplicate independent RTL example', await page.getByRole('button', { name: /Independent RTL instance/ }).count(), 0);
  // Independent instance/provider behavior is exercised by dedicated card/Stage fixtures.
  await capture(page, 'live-custom-retention');
  await select(page, 'stage');
  await page.locator('[data-stage-demo=trigger]').first().click();
  const dialog = page.getByRole('dialog').first(); await dialog.waitFor();
  await page.evaluate(() => window.demoFixtureHost.update({ theme: 'custom', brandColor: '#7b2cbf', dir: 'rtl' })); await settle(page);
  check('Popup live custom token', await dialog.evaluate(element => getComputedStyle(element).getPropertyValue('--colorBrandBackground').trim()), '#7b2cbf');
  check('Popup focus owned inside surface', await dialog.evaluate(element => element.contains(document.activeElement)), true);
  await page.waitForFunction(() => { const element = document.querySelector('[role=dialog]'); return element && getComputedStyle(element).opacity === '1' && element.getAnimations({ subtree: true }).every(animation => animation.playState !== 'running'); });
  check('Settled popup computed RTL direction', await dialog.evaluate(element => getComputedStyle(element).direction), 'rtl');
  await capture(page, 'popup-custom-rtl'); await page.keyboard.press('Escape'); await dialog.waitFor({ state: 'hidden' });
  check('Popup closes with Escape', await dialog.count(), 0);
  check('Popup restores trigger focus', await page.locator('[data-stage-demo=trigger]').first().evaluate(element => document.activeElement === element), true);
}
(async () => {
  fs.mkdirSync(output, { recursive: true }); await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const url = `http://127.0.0.1:${server.address().port}`; let browser;
  try {
    browser = await chromium.launch({ headless: true, executablePath });
    const page = await browser.newPage({ viewport: { width: 1280, height: 1000 } }); observe(page); page.setDefaultTimeout(10000);
    await guardNetwork(page, url);
    await page.goto(url); check('Page identity', await page.title(), 'Integrated demo presentation checkpoint');
    check('Local origin', new URL(page.url()).origin, url);
    check('Framework overlays absent', await page.locator('webpack-dev-server-client-overlay,#webpack-dev-server-client-overlay,nextjs-portal').count(), 0);
    check('Actual registered tab count', await page.getByRole('tab').count(), entries.length);
    check('Coverage detailed default collapsed', await page.getByRole('button', { name: 'Detailed package coverage', exact: true }).getAttribute('aria-expanded'), 'false');
    await page.screenshot({ path: path.join(output, 'initial-viewport.png') });
    if (process.env.DEMO_PRESENTATION_TRIMS_ONLY === '1') {
      const { accessTrimsBrowser } = require('./access-trims-browser.helpers.cjs');
      await accessTrimsBrowser({ page, check, select, singleExample, labels, overflow, capture });
      check('Unexpected trim browser/network diagnostics', diagnostics, []);
      console.log(`Actual ${process.env.DEMO_PRESENTATION_MODE || 'compiled'} trim panels and scoped hook: ${checks.length} checks PASS.`);
      return;
    }
    if (process.env.DEMO_PRESENTATION_GROUPS_ONLY === '1') {
      const restoredCapture = async name => {
        await page.evaluate(async () => {
          await Promise.all(document.getAnimations().filter(animation => animation.effect?.getTiming().iterations !== Infinity).map(animation => animation.finished.catch(() => undefined)));
        });
        await settle(page);
        await capture(page, name);
        const subjectFile = path.join(output, `${name}-subject.png`);
        await page.getByRole('tabpanel').locator('[data-demo-subject]').screenshot({ path: subjectFile }); screenshots.push(subjectFile);
      };
      await select(page, 'groups');
      check('Groups visible card appearance is the demo default', await page.locator('select[data-groups-demo=appearance]').inputValue(), 'filled');
      check('Groups appearance has only approved values', await page.locator('select[data-groups-demo=appearance] option').evaluateAll(nodes => nodes.map(node => node.value)), ['filled', 'alternate', 'outline']);
      await groupsProduct(page, { check });
      const panel = page.getByRole('tabpanel');
      for (const type of ['settings', 'accordion']) {
        await panel.locator('select[data-demo-variant=GroupType]').selectOption(type);
        if (type === 'accordion') {
          await panel.locator('[data-groups-demo=close-all]').click();
          await panel.getByRole('button', { name: /^Profile details\s*Your account preferences\.$/ }).click();
        }
        for (const rtl of [false, true]) {
          await panel.locator('input[data-groups-demo=direction]').setChecked(rtl);
          for (const width of [1280, 320]) {
            await page.setViewportSize({ width, height: 1000 }); await settle(page);
            await singleExample(page, 'groups', `${type}/${rtl ? 'rtl' : 'ltr'}/${width}`);
            await overflow(page, `groups/${type}/${rtl ? 'rtl' : 'ltr'}/${width}`);
            const group = panel.locator(`[data-groups-demo=${type}]`);
            const bodyCard = group.locator(':scope > .fui-Card').last();
            const input = group.locator('input').first();
            const retainedInput = await input.elementHandle();
            const retainedValue = await input.inputValue();
            const surfaces = [];
            for (const appearance of ['filled', 'alternate', 'outline']) {
              await panel.locator('select[data-groups-demo=appearance]').selectOption(appearance);
              await settle(page);
              await singleExample(page, 'groups', `${type}/${appearance}/${rtl ? 'rtl' : 'ltr'}/${width}`);
              await overflow(page, `groups/${type}/${appearance}/${rtl ? 'rtl' : 'ltr'}/${width}`);
              const surface = await bodyCard.evaluate(element => {
                const css = getComputedStyle(element);
                const border = getComputedStyle(element, '::after');
                return { background: css.backgroundColor, shadow: css.boxShadow, border: border.borderTopStyle, borderWidth: border.borderTopWidth };
              });
              surfaces.push(surface);
              check(`groups ${type} ${appearance} ${width} native surface`, appearance === 'outline'
                ? surface.background === 'rgba(0, 0, 0, 0)' && surface.shadow === 'none' && surface.border === 'solid' && parseFloat(surface.borderWidth) > 0
                : surface.background !== 'rgba(0, 0, 0, 0)' && surface.shadow !== 'none', true);
              check(`groups ${type} ${appearance} retains edited input`, await retainedInput.evaluate((element, value) => element.isConnected && element.value === value, retainedValue), true);
              await restoredCapture(`groups-${type}-${appearance}-${rtl ? 'rtl' : 'ltr'}-${width}`);
            }
            check(`groups ${type} ${width} alternate is a distinct native color`, surfaces[0].background !== surfaces[1].background, true);
            await retainedInput.dispose();
            await panel.locator('select[data-groups-demo=appearance]').selectOption('filled');
            check(`groups ${type} ${width} direction`, await group.evaluate(element => getComputedStyle(element).direction), rtl ? 'rtl' : 'ltr');
            const icons = await group.locator('[data-group-item-icon]').evaluateAll(nodes => nodes.map(node => Math.round(node.getBoundingClientRect().width)));
            check(`groups ${type} ${width} bounded item icons`, icons, [20, 20]);
            await restoredCapture(`groups-${type}-${rtl ? 'rtl' : 'ltr'}-${width}`);
          }
        }
      }
      await page.setViewportSize({ width: 1280, height: 1000 });
      await select(page, 'auto-skeleton');
      const auto = id => panel.locator(`[data-auto-skeleton-demo=${id}]`);
      await auto('card').waitFor();
      check('AutoSkeleton profile has one native Fluent card', await auto('card').count(), 1);
      check('AutoSkeleton profile is inside the native card', await auto('card').evaluate(element => element.classList.contains('fui-Card') && Boolean(element.querySelector('[data-auto-skeleton-demo=preview]'))), true);
      check('AutoSkeleton actual named avatar', await auto('circle').getAttribute('aria-label'), 'Giulia Rossi');
      check('AutoSkeleton rectangle contains actual profile title', await auto('rectangle').textContent(), 'Giulia Rossi');
      check('AutoSkeleton actual profile role', await panel.getByText('Project coordinator', { exact: true }).count(), 1);
      await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click();
      const note = auto('focus-target'); await note.fill('Retained profile draft'); const retained = await note.elementHandle();
      for (const rtl of [false, true]) {
        await auto('direction').setChecked(rtl);
        for (const width of [1280, 320]) {
          await page.setViewportSize({ width, height: 1000 }); await auto('resize').setChecked(width === 320); await settle(page);
          await singleExample(page, 'auto-skeleton', `${rtl ? 'rtl' : 'ltr'}/${width}`);
          await overflow(page, `auto-skeleton/${rtl ? 'rtl' : 'ltr'}/${width}`);
          check('AutoSkeleton card uses a visible themed surface', await auto('card').evaluate(element => getComputedStyle(element).boxShadow !== 'none' && getComputedStyle(element).backgroundColor !== 'rgba(0, 0, 0, 0)'), true);
          await restoredCapture(`auto-skeleton-loaded-${rtl ? 'rtl' : 'ltr'}-${width}`);
          await auto('focus-load').click(); await auto('preview').locator('[data-auto-skeleton-overlay]').waitFor();
          check('AutoSkeleton card surface stays visible during loading', await auto('card').evaluate(element => getComputedStyle(element).visibility), 'visible');
          check('AutoSkeleton loading excludes content from focus', await note.evaluate(element => !element.closest('[data-auto-skeleton-content]').contains(document.activeElement)), true);
          check('AutoSkeleton actual circle placeholder exists', await auto('preview').locator('[data-auto-skeleton-shape=circle]').count() > 0, true);
          check('AutoSkeleton actual rectangle placeholder exists', await auto('preview').locator('[data-auto-skeleton-shape=rectangle]').count() > 0, true);
          check('AutoSkeleton has no square placeholders', await auto('preview').locator('[data-auto-skeleton-shape=square]').count(), 0);
          await restoredCapture(`auto-skeleton-loading-${rtl ? 'rtl' : 'ltr'}-${width}`);
          await auto('loading').uncheck(); await auto('preview').locator('[data-auto-skeleton-overlay]').waitFor({ state: 'detached' });
          check('AutoSkeleton preserves note node and value', await retained.evaluate(element => element === document.querySelector('[data-auto-skeleton-demo=focus-target]') && element.value === 'Retained profile draft'), true);
          check('AutoSkeleton child is not remounted on presentation changes', (await panel.textContent()).includes('Content mounts: 1'), true);
        }
      }
      await retained.dispose();
      check('Groups and AutoSkeleton unexpected diagnostics', diagnostics, []);
      console.log(`Groups and AutoSkeleton restoration: ${checks.length} checks PASS.`);
      return;
    }
    if (process.env.DEMO_PRESENTATION_ENGLISH_ONLY === '1') {
      const englishCopy = async key => {
        const panel = page.getByRole('tabpanel');
        check(`${key}: English configuration heading`, await panel.getByText(/^Configuration and (simulation|results)$/).count(), 1);
        const diagnosticHeaders = panel.locator('[data-demo-configuration] > .fui-Accordion > .fui-AccordionItem > .fui-AccordionHeader button');
        check(`${key}: English diagnostics label when present`, await panel.getByRole('button', { name: 'Advanced diagnostics', exact: true }).count(), await diagnosticHeaders.count());
        const copy = await panel.evaluate(element => [element.textContent, ...[...element.querySelectorAll('input,textarea,[aria-label],[title]')].flatMap(node => [node.value || '', node.getAttribute('aria-label') || '', node.getAttribute('title') || ''])].join(' '));
        check(`${key}: old Italian captions absent`, /\b(?:Configurazione|Diagnostica|Titolo|Descrizione|Larghezza|Caricamento|Esegui|Salva|Ripristina|Operazione|Ordine|Ordini|Preferenze|Contatti|Aggiornamento|Immagine|Contenuto|Proprietà|Nessuna|Nessun)\b/u.test(copy), false);
      };
      for (const entry of entries) {
        await select(page, entry.key);
        await englishCopy(entry.key);
        await singleExampleTransitions(page, entry.key, () => englishCopy(entry.key));
      }
      await select(page, 'cards');
      const title = page.locator('[data-card-instance=main] [data-cards-demo=title-input]');
      check('Cards English default title', await title.inputValue(), 'Quarterly report');
      await page.getByRole('tabpanel').screenshot({ path: path.join(output, 'cards-english.png') });
      await select(page, 'relative-time');
      check('RelativeTime English default locale', await page.locator('[data-relative-time-demo=locale]').inputValue(), 'en-US');
      await interact(page, 'relative-time');
      await select(page, 'pagination');
      check('Pagination English custom labels', await page.getByRole('button', { name: 'Next page', exact: true }).count(), 1);
      const nativeColorWarning = "The keyword 'slider-vertical' specified to an 'appearance' property is not standardized. It will be removed in the future. Use <input type=range style=\"writing-mode: vertical-lr; direction: rtl\"> instead.";
      check('Unexpected English smoke diagnostics', diagnostics.filter(message => !(message.type === 'warning' && message.text === nativeColorWarning)), []);
      console.log(`English demo copy: ${entries.length} tabs and selectable variants PASS (${checks.length} checks).`);
      return;
    }
    if (process.env.DEMO_PRESENTATION_CARDINALITY_ONLY === '1') {
      for (const entry of entries) { await select(page, entry.key); await singleExampleTransitions(page, entry.key); }
      console.log(`Single-example cardinality: ${entries.length} tabs PASS`);
      return;
    }
    for (const entry of entries) {
      await select(page, entry.key); check(`${entry.key}: meaningful primary example`, await page.getByRole('tabpanel').locator('h3,[data-shell-demo=panel]').count() > 0, true);
      await singleExampleTransitions(page, entry.key);
      check(`${entry.key}: duplicate RTL demos absent`, await page.getByRole('tabpanel').getByRole('button', { name: /Independent RTL instance/ }).count(), 0);
      const layoutIssues = await page.getByRole('tabpanel').evaluate(panel => [...panel.querySelectorAll('[data-demo-section]')].flatMap(section => {
        const tools = section.querySelector(':scope > [data-demo-configuration]');
        const issues = [];
        if (!tools || getComputedStyle(tools).backgroundColor !== 'rgba(0, 0, 0, 0)' || getComputedStyle(tools).boxShadow !== 'none') issues.push('Configuration card must be transparent and have no shadow');
        if (tools?.querySelector('h3')) issues.push('Heading must be outside configuration');
        for (const subject of section.querySelectorAll(':scope > [role=group]:not([data-demo-configuration])')) {
          const style = getComputedStyle(subject);
          if (style.backgroundColor !== 'rgba(0, 0, 0, 0)' || ['paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft'].some(name => style[name] !== '0px')) issues.push(`Live subject wrapper must have no background or padding: ${subject.getAttribute('aria-label')}, ${style.backgroundColor}, ${style.padding}`);
        }
        return issues;
      }));
      check(`${entry.key}: approved section structure`, layoutIssues, []);
      await labels(page, entry.key);
      try { await interact(page, entry.key); } catch (error) {
        fs.writeFileSync(path.join(output, 'interaction-debug.json'), JSON.stringify(await page.getByRole('tabpanel').evaluate(panel => ({ text: panel.textContent?.slice(0, 6000), controls: [...panel.querySelectorAll('input,select,button')].map(element => ({ tag: element.tagName, label: element.getAttribute('aria-label'), text: element.tagName === 'BUTTON' ? element.textContent : undefined, value: element.value })) })), undefined, 2));
        throw error;
      }
      await singleExample(page, entry.key, 'after-interaction'); console.log(`Verified interaction: ${entry.key}`);
      for (const width of [1280, 640, 320]) { await page.setViewportSize({ width, height: 1000 }); await settle(page); await singleExampleTransitions(page, entry.key, variant => overflow(page, `${entry.key}/${width}/${variant}`)); await capture(page, `${entry.key}-${width}`); await overflow(page, `${entry.key}/${width}`); }
      await page.setViewportSize({ width: 1280, height: 1000 });
    }
    await liveTheme(page);
    await page.evaluate(() => window.demoFixtureHost.update({ theme: 'light', dir: 'ltr' }));
    await page.setViewportSize({ width: 320, height: 1000 }); await settle(page);
    const firstTab = page.getByRole('tab').first(); await firstTab.focus(); await page.keyboard.press('ArrowRight'); await settle(page);
    check('Keyboard arrow reaches and activates second tab', await page.getByRole('tab').nth(1).getAttribute('aria-selected'), 'true');
    const tabs = page.getByRole('tab');
    for (let i = 2; i < entries.length; i++) { await page.keyboard.press('ArrowRight'); await settle(page); check(`Keyboard reaches tab ${i + 1}`, await tabs.nth(i).getAttribute('aria-selected'), 'true'); }
    await capture(page, 'keyboard-navigation-320');
    for (const zoom of [1, 2]) {
      const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-presentation-native-zoom-')); fs.mkdirSync(path.join(profile, 'Default'));
      fs.writeFileSync(path.join(profile, 'Default', 'Preferences'), JSON.stringify({ partition: { default_zoom_level: { x: Math.log(zoom) / Math.log(1.2) } } }));
      let context;
      try {
        context = await chromium.launchPersistentContext(profile, { executablePath, headless: true, viewport: null, args: ['--window-size=1440,1000'] });
        const native = context.pages()[0]; observe(native); await guardNetwork(native, url); await native.goto(url); await native.getByRole('tabpanel').locator('h3').first().waitFor();
        metrics.push(await native.evaluate(() => ({ width: innerWidth, height: innerHeight, dpr: devicePixelRatio, scale: visualViewport.scale })));
        for (const { key } of entries) { await select(native, key); await singleExampleTransitions(native, key, variant => overflow(native, `${key}/native-${zoom * 100}%/${variant}`)); await overflow(native, `${key}/native-${zoom * 100}%`); await labels(native, `${key}/native-${zoom * 100}%`); await native.evaluate(() => document.querySelector('[role=tabpanel]').scrollIntoView({ block: 'start' }));
          const session = await context.newCDPSession(native);
          const shot = await session.send('Page.captureScreenshot', { format: 'png', fromSurface: true, captureBeyondViewport: false });
          const buffer = Buffer.from(shot.data, 'base64');
          const filename = path.join(output, `${key}-native-${zoom * 100}-viewport-cdp.png`); fs.writeFileSync(filename, buffer); screenshots.push(filename);
          check(`${key}/native-${zoom * 100}% physical capture width`, buffer.readUInt32BE(16), 1440); await session.detach(); }
      } finally { await context?.close(); fs.rmSync(profile, { recursive: true, force: true }); }
    }
    check('Native 200% doubles DPR', metrics[1].dpr / metrics[0].dpr, 2);
    check('Native 200% halves CSS viewport', metrics[0].width / metrics[1].width, 2);
    const nativeColorWarning = "The keyword 'slider-vertical' specified to an 'appearance' property is not standardized. It will be removed in the future. Use <input type=range style=\"writing-mode: vertical-lr; direction: rtl\"> instead.";
    check('Unexpected app/browser/network diagnostics', diagnostics.filter(message => !(message.type === 'warning' && message.text === nativeColorWarning)), []);
    console.log(`Integrated actual app: ${checks.length} PASS checks; ${screenshots.length} screenshots; native 200% proved.`);
  } finally {
    fs.writeFileSync(path.join(output, 'browser-results.json'), JSON.stringify({ browser: browser?.version(), browserPlugin: 'not available', fallback: 'Bundled Playwright + installed Chrome', boundary: `actual app/registry/lazy panels + ${process.env.DEMO_PRESENTATION_MODE || 'compiled'} controls/styles/hooks and real provider; native SP/Graph transport doubles; no tenant CRUD`, sharePointHost: 'NOT EXECUTED', unavailableFontBoundary: 'Exact legacy CDN Segoe assets fulfilled with installed Arial; no external requests', fontAssets, checks, diagnostics, metrics, screenshots }, undefined, 2));
    await browser?.close(); await new Promise(resolve => server.close(resolve));
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
