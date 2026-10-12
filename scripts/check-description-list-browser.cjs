// Run the actual sample through isolated source and compiled library boundaries.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const root = path.resolve(__dirname, '..');
const evidence = path.resolve(process.env.DESCRIPTION_LIST_BROWSER_EVIDENCE || path.join(root, '.docs/maintenance/evidence/description-list'));

async function verify(boundary) {
  const directory = path.join(root, 'temp', boundary === 'source' ? 'sx-browser' : 'styles-panel-browser');
  const server = http.createServer((request, response) => {
    const pathname = new URL(request.url, 'http://127.0.0.1').pathname;
    if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
    const file = { '/': 'index.html', '/fixture.js': 'fixture.js', '/fixture.js.map': 'fixture.js.map' }[pathname];
    if (!file) { response.writeHead(404).end(); return; }
    response.setHeader('Content-Type', file.endsWith('.html') ? 'text/html' : 'application/javascript');
    response.end(fs.readFileSync(path.join(directory, file)));
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  let browser;
  const checks = [], errors = [], warnings = [];
  const check = (name, value, expected) => { assert.deepEqual(value, expected, name); checks.push(name); };
  const geometry = list => list.evaluate(node => {
    const terms = [...node.querySelectorAll('dt')], definitions = [...node.querySelectorAll('dd')];
    const term = terms[0]?.getBoundingClientRect(), value = definitions[0]?.getBoundingClientRect();
    return {
      horizontal: !!term && !!value && Math.abs(term.top - value.top) < 2,
      aligned: definitions.every(cell => Math.abs(cell.getBoundingClientRect().left - value.left) < 2),
      overflow: node.scrollWidth > node.clientWidth + 1,
      inlineStyles: node.hasAttribute('style') || [...node.querySelectorAll('dt,dd')].some(cell => cell.hasAttribute('style')),
      terms: terms.length, definitions: definitions.length,
    };
  });
  try {
    browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
    const page = await browser.newPage({ viewport: { width: 1280, height: 1100 } });
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); if (message.type() === 'warning') warnings.push(message.text()); });
    const url = `http://127.0.0.1:${server.address().port}`;
    await page.goto(url);
    const primary = page.locator('[data-description-list-instance="primary"]');
    const secondary = page.locator('[data-description-list-instance="secondary"]');
    const list = primary.locator('[data-description-list-demo="list"]');
    const select = name => primary.locator(`[data-description-list-demo="${name}"]`);
    const checkbox = label => primary.getByRole('checkbox', { name: label, exact: true });
    check('title', await page.title(), `DescriptionList ${boundary} browser checkpoint`);
    check('boundary', await page.locator('[data-description-list-fixture="boundary"]').textContent(), boundary);
    check('page identity', new URL(page.url()).origin, url);
    check('meaningful nonblank sample', await page.getByRole('heading', { name: 'Description List', exact: true }).isVisible(), true);
    check('no framework overlay', await page.locator('webpack-dev-server-client-overlay, nextjs-portal').count(), 0);
    check('native semantic root', await list.evaluate(node => node.tagName), 'DL');
    check('data pair count', (await geometry(list)).terms, 10);
    check('native group structure', await list.evaluate(node => [...node.children].every(group => group.tagName === 'DIV' && group.children.length === 2 && group.children[0].tagName === 'DT' && group.children[1].tagName === 'DD')), true);
    check('literal content', await list.locator('dt').nth(2).textContent(), '**Literal label**');
    check('zero preserved', await list.locator('dd').nth(3).textContent(), '0');
    check('blank preserved', await list.locator('dd').nth(4).textContent(), '');
    check('undefined empty definition', await list.locator('dd').nth(5).textContent(), '');
    check('label native slot', await list.locator('dt [title="Native label slot title"]').count(), 1);
    check('callback slot block content', await list.locator('dd p').textContent(), 'Block content from a slot callback.');
    await list.getByRole('link', { name: 'Read details' }).click();
    check('callback retains native link', new URL(page.url()).hash, '#description-list-demo');
    check('click handler bubbles', await primary.locator('[data-description-list-demo="clicks"]').textContent(), 'Bubbled list clicks: 1');
    for (const mode of ['auto', 'horizontal', 'vertical']) {
      await select('layout').selectOption(mode);
      for (const pixels of [320, 479, 480, 639, 640, 900]) {
        await select('width').selectOption(String(pixels));
        await page.waitForFunction(({ instance, pixels }) => {
          const node = document.querySelector(`[data-description-list-instance="${instance}"] [data-description-list-demo="container"]`);
          return node && Math.abs(node.getBoundingClientRect().width - pixels) < 1;
        }, { instance: 'primary', pixels });
        await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
        const box = await geometry(list);
        check(`${mode} ${pixels}px orientation`, box.horizontal, mode === 'horizontal' || mode === 'auto' && pixels >= 480);
        check(`${mode} ${pixels}px pair invariant`, [box.terms, box.definitions], [10, 10]);
        check(`${mode} ${pixels}px long content fits`, box.overflow, false);
        check(`${mode} ${pixels}px styles use classes`, box.inlineStyles, false);
        if (box.horizontal) check(`${mode} ${pixels}px shared definition column`, box.aligned, true);
        if (mode === 'auto') await list.screenshot({ path: path.join(evidence, `${boundary}-auto-${pixels}.png`) });
      }
    }
    await select('layout').selectOption('auto');
    await select('width').selectOption('900');
    check('nearest nested ancestor', (await geometry(primary.locator('[data-description-list-demo="nested-container"] dl'))).horizontal, false);
    const fallback = page.locator('[data-description-list-probe="fallback"]');
    for (const viewport of [320, 900, 1280]) {
      await page.setViewportSize({ width: viewport, height: 1100 });
      check(`no named ancestor ${viewport}px fallback`, (await geometry(fallback)).horizontal, false);
      check(`900px named ancestor wins over ${viewport}px viewport`, (await geometry(list)).horizontal, true);
    }
    const counter = primary.getByRole('button', { name: 'primary A: 0', exact: true });
    await counter.click();
    await select('width').selectOption('320');
    await select('layout').selectOption('vertical');
    await checkbox('Reverse stable IDs').check();
    check('caller state retained across resize layout reorder', await primary.getByRole('button', { name: 'primary A: 1', exact: true }).count(), 1);
    check('independent instance state', await secondary.getByRole('button', { name: 'secondary A: 0', exact: true }).count(), 1);
    check('independent instance layout', await secondary.locator('[data-description-list-demo="layout"]').inputValue(), 'auto');
    const input = page.getByRole('textbox', { name: 'Retained value', exact: true });
    await input.fill('caller retained value');
    await input.focus();
    await input.evaluate(node => { window.descriptionListRetainedInput = node; });
    for (const action of ['resize', 'layout', 'reverse']) {
      await page.locator(`[data-description-list-probe="${action}"]`).evaluate(node => node.click());
      check(`input focus retained after ${action}`, await input.evaluate(node => document.activeElement === node), true);
      check(`input node retained after ${action}`, await input.evaluate(node => window.descriptionListRetainedInput === node), true);
      check(`input value retained after ${action}`, await input.inputValue(), 'caller retained value');
    }
    await checkbox('Reverse stable IDs').uncheck();
    await checkbox('Compound JSX').check();
    check('compound pair count', (await geometry(list)).terms, 10);
    await primary.locator('[data-description-list-demo="inspect-ref"]').click();
    check('native root and item refs', (await primary.locator('[data-description-list-demo="ref"]').textContent()).includes('root=DL; pairs=10; root hidden=false; item=DIV; item hidden=false'), true);
    await checkbox('Hide first compound pair').check();
    check('native item hidden', await list.locator('[data-description-list-demo="pair"]').first().isVisible(), false);
    await checkbox('Hide root').check();
    check('native root hidden', await list.isVisible(), false);
    await primary.locator('[data-description-list-demo="inspect-ref"]').click();
    check('hidden ref still native', (await primary.locator('[data-description-list-demo="ref"]').textContent()).includes('root hidden=true; item=DIV; item hidden=true'), true);
    await checkbox('Hide root').uncheck();
    await checkbox('Hide first compound pair').uncheck();
    await checkbox('Empty input').check();
    check('empty compound root survives', await list.count(), 1);
    check('empty compound pairs', await list.locator('dt,dd').count(), 0);
    await checkbox('Compound JSX').uncheck();
    check('empty data pairs', await list.locator('dt,dd').count(), 0);
    check('empty data renders nothing', await list.count(), 0);
    await checkbox('Empty input').uncheck();
    await select('width').selectOption('640');
    await select('layout').selectOption('auto');
    await checkbox('Dark theme').check();
    await checkbox('RTL').check();
    check('RTL inherited', await list.evaluate(node => getComputedStyle(node).direction), 'rtl');
    check('RTL horizontal shared columns', (await geometry(list)).aligned, true);
    check('dark visible values', await list.getByRole('link', { name: 'Ada Lovelace' }).isVisible(), true);
    await primary.screenshot({ path: path.join(evidence, `${boundary}-dark-rtl.png`) });
    await page.emulateMedia({ forcedColors: 'active' });
    check('forced colors preserves semantic content', (await geometry(list)).terms, 10);
    check('forced colors fits', (await geometry(list)).overflow, false);
    await primary.screenshot({ path: path.join(evidence, `${boundary}-forced-colors.png`) });
    await page.emulateMedia({ forcedColors: 'none' });
    await checkbox('Dark theme').uncheck();
    await checkbox('RTL').uncheck();
    await primary.screenshot({ path: path.join(evidence, `${boundary}-light.png`) });
    await select('width').selectOption('320');
    await page.setViewportSize({ width: 320, height: 1100 });
    await list.screenshot({ path: path.join(evidence, `${boundary}-320.png`) });
    check('mobile list fits', (await geometry(list)).overflow, false);
    check('console errors', errors, []);
    check('console warnings', warnings, []);
    return { boundary, url, checks, errors, warnings, browser: browser.version(), containerWidths: [320, 479, 480, 639, 640, 900] };
  } finally { await browser?.close(); await new Promise(resolve => server.close(resolve)); }
}
async function main() {
  fs.mkdirSync(evidence, { recursive: true });
  const reports = [];
  for (const boundary of ['source', 'compiled']) { const report = await verify(boundary); reports.push(report); console.log(`DescriptionList ${boundary}: ${report.checks.length} browser checks passed (${report.browser}).`); }
  fs.writeFileSync(path.join(evidence, 'browser-results.json'), JSON.stringify({ browserPlugin: 'Browser plugin not available', fallback: 'Existing Playwright module and installed Chrome; no dependencies installed', reports }, undefined, 2));
}
main().catch(error => { console.error(error); process.exitCode = 1; });
