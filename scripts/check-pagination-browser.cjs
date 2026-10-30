const { paginationProduct } = require('../tests/fixtures/demo-control-isolation/task3-product.browser.cjs');
// Exercise the actual sample against both source and compiled library boundaries.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const { teamsDarkTheme } = require('@fluentui/react-theme');
const root = path.resolve(__dirname, '..');
const evidence = path.resolve(process.env.PAGINATION_BROWSER_EVIDENCE || path.join(root, '.docs/maintenance/evidence/pagination'));

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
  const url = `http://127.0.0.1:${server.address().port}`;
  let browser;
  const checks = [], errors = [], warnings = [];
  const check = (name, value, expected) => { assert.deepEqual(value, expected, name); checks.push(name); };
  try {
    browser = await chromium.launch({ headless: true,
      executablePath: process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
    const page = await browser.newPage({ viewport: { width: 1440, height: 1100 } });
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => {
      if (message.type() === 'error') errors.push(message.text());
      if (message.type() === 'warning') warnings.push(message.text());
    });
    await paginationProduct(page, { url, check });
    await page.goto(url);
    await page.getByRole('button', { name: 'Independent RTL instance', exact: true }).click();
    const primary = page.locator('[data-pagination-instance="primary"]');
    const secondary = page.locator('[data-pagination-instance="secondary"]');
    const diagnostics = primary.getByRole('button', { name: 'Advanced diagnostics', exact: true });
    if (await diagnostics.getAttribute('aria-expanded') !== 'true') await diagnostics.click();
    const summary = primary.locator('[data-pagination-demo="summary"]');
    const nav = primary.locator('nav');
    const action = label => nav.getByRole('button', { name: label, exact: true });
    const select = name => primary.locator(`[data-pagination-demo="${name}"]`);
    const checkbox = label => primary.getByRole('checkbox', { name: label, exact: true });
    check('title', await page.title(), `Pagination ${boundary} browser checkpoint`);
    check('boundary', await page.locator('[data-pagination-fixture="boundary"]').textContent(), boundary);
    check('page identity', new URL(page.url()).origin, url);
    check('meaningful first page', (await summary.textContent()).trim(), 'Page 1: items 1–5 of 23.');
    check('no framework overlay', await page.locator('webpack-dev-server-client-overlay, nextjs-portal').count(), 0);
    check('native size options', await select('size').locator('option').evaluateAll(options => options.map(option => option.value)), ['small', 'medium', 'large']);
    check('native appearance options', await select('appearance').locator('option').evaluateAll(options => options.map(option => option.value)), ['primary', 'secondary', 'outline', 'subtle', 'transparent']);
    check('default subtle appearance', await select('appearance').inputValue(), 'subtle');
    check('first unavailable', await action('First page').isDisabled(), true);
    check('native non-submit buttons', await nav.locator('button').evaluateAll(nodes => nodes.every(node => node.type === 'button')), true);
    await action('Next page').focus();
    await page.keyboard.press('Enter');
    check('keyboard commits next page', (await summary.textContent()).trim(), 'Page 2: items 6–10 of 23.');
    check('actual row changed', await primary.getByRole('gridcell', { name: 'Order 06', exact: true }).count(), 1);
    check('second instance independent', (await secondary.locator('[data-pagination-demo="summary"]').textContent()).trim(), 'Page 1: items 1–5 of 23.');
    const appearanceClasses = [];
    const subtleClass = await action('Next page').getAttribute('class');
    for (const appearance of ['primary', 'secondary', 'outline', 'subtle', 'transparent']) {
      await select('appearance').selectOption(appearance);
      const buttonClass = await action('Next page').getAttribute('class');
      appearanceClasses.push(buttonClass);
      check(`${appearance} applies to every page action`, await nav.locator('button').evaluateAll(nodes => nodes.every(node => node.className === nodes[0].className)), true);
      check(`${appearance} does not change page results`, (await summary.textContent()).trim(), 'Page 2: items 6–10 of 23.');
    }
    check('five native appearance presentations', new Set(appearanceClasses).size, 5);
    check('appearance instance isolation', await secondary.locator('[data-pagination-demo="appearance"]').inputValue(), 'subtle');
    check('appearance belongs to actions', await nav.getAttribute('appearance'), null);
    await select('appearance').selectOption('subtle');
    check('subtle appearance restored', await action('Next page').getAttribute('class'), subtleClass);
    await checkbox('Complete requests manually').check();
    // Finish disclosure layout and Playwright's scrolling before comparing icon geometry.
    await page.evaluate(() => Promise.all(document.getAnimations().filter(animation => animation.effect?.getTiming().iterations !== Infinity).map(animation => animation.finished.catch(() => undefined))));
    await action('Next page').scrollIntoViewIfNeeded();
    const before = await action('Next page').boundingBox();
    await action('Next page').click();
    const after = await action('Next page').boundingBox();
    check('stable action icon position', { x: after.x, y: after.y, width: after.width }, { x: before.x, y: before.y, width: before.width });
    check('pending keeps visible rows', await primary.getByRole('gridcell', { name: 'Order 06', exact: true }).count(), 1);
    check('pending locks all actions', await nav.locator('button').evaluateAll(nodes => nodes.every(node => node.disabled)), true);
    check('loading accessible', await nav.getAttribute('aria-busy'), 'true');
    check('progress labelled', await nav.getByRole('progressbar', { name: 'Loading', exact: true }).count(), 1);
    await primary.locator('[data-pagination-demo="complete"]').click();
    check('completion commits once', (await summary.textContent()).trim(), 'Page 3: items 11–15 of 23.');
    check('compact center gap', await nav.evaluate(node => { const buttons = node.querySelectorAll('button'); const previous = buttons[1].getBoundingClientRect(), next = buttons[2].getBoundingClientRect(); return Math.round(next.left - previous.right); }), 4);
    await action('Previous page').click();
    for (const [size, diameter, buttonWidth] of [['small', 16, 24], ['medium', 20, 32], ['large', 24, 40]]) {
      await select('size').selectOption(size);
      check(`${size} progress diameter`, await action('Previous page').locator('.fui-Spinner__spinner').evaluate(node => { return { width: node.offsetWidth, height: node.offsetHeight }; }), { width: diameter, height: diameter });
      check(`${size} native icon-button width`, (await action('Previous page').boundingBox()).width, buttonWidth);
    }
    await select('size').selectOption('medium');
    check('spinner replaces requested previous icon', await action('Previous page').getByRole('progressbar', { name: 'Loading', exact: true }).count(), 1);
    check('other direction keeps arrow', await action('Next page').locator('svg').count(), 1);
    const defaultPendingClass = await action('Previous page').getAttribute('class');
    await select('appearance').selectOption('primary');
    check('pending appearance changes presentation', await action('Previous page').getAttribute('class') !== defaultPendingClass, true);
    check('pending appearance keeps proportional spinner', await action('Previous page').locator('.fui-Spinner__spinner').evaluate(node => node.offsetWidth), 20);
    await select('appearance').selectOption('subtle');
    await primary.screenshot({ path: path.join(evidence, `${boundary}-loading.png`) });
    await primary.locator('[data-pagination-demo="complete"]').click();
    check('previous completion commits once', (await summary.textContent()).trim(), 'Page 2: items 6–10 of 23.');
    await checkbox('Complete requests manually').uncheck();
    await select('page-size').selectOption('3');
    check('size reset', (await summary.textContent()).trim(), 'Page 1: items 1–3 of 23.');
    await action('Last page').click();
    check('last page interval', (await summary.textContent()).trim(), 'Page 8: items 22–23 of 23.');
    check('last unavailable', await action('Next page').isDisabled(), true);
    await select('filter').selectOption('Hardware');
    check('filter reset', (await summary.textContent()).trim(), 'Page 1: items 1–3 of 12.');
    await select('filter').selectOption('all');
    await select('mode').selectOption('unknown');
    check('unknown total hides extremes', await nav.locator('button').count(), 2);
    await action('Next page').click();
    check('unknown offset request', (await primary.locator('[data-pagination-demo="request"]').textContent()).trim(), 'Last request: {"type":"pages","pageIndex":1,"pageSize":3}');
    await select('mode').selectOption('cursor');
    await action('Next page').click();
    check('cursor direction request', (await primary.locator('[data-pagination-demo="request"]').textContent()).trim(), 'Last request: {"type":"cursor","direction":"next"}');
    await action('Previous page').click();
    check('cursor previous results', (await summary.textContent()).trim(), 'Page 1: items 1–3 of 23.');
    await select('mode').selectOption('loadMore');
    check('load-more has no idle icon', await action('Load more').locator('.fui-Button__icon').count(), 0);
    await select('appearance').selectOption('outline');
    check('load-more appearance is selectable', await select('appearance').inputValue(), 'outline');
    check('load-more appearance keeps text-only action', await action('Load more').locator('.fui-Button__icon').count(), 0);
    await primary.screenshot({ path: path.join(evidence, `${boundary}-load-more-idle.png`) });
    await action('Load more').click();
    check('incremental results', (await summary.textContent()).trim(), 'Loaded items 1–6 of 23.');
    check('incremental request', (await primary.locator('[data-pagination-demo="request"]').textContent()).trim(), 'Last request: {"type":"loadMore"}');
    await checkbox('Complete requests manually').check();
    const loadBefore = await action('Load more').boundingBox();
    await action('Load more').click();
    const loadAfter = await action('Load more').boundingBox();
    check('load-more remains centered', Math.round(loadAfter.x + loadAfter.width / 2), Math.round(loadBefore.x + loadBefore.width / 2));
    for (const [size, diameter] of [['small', 16], ['medium', 20], ['large', 24]]) {
      await select('size').selectOption(size);
      check(`load-more ${size} progress diameter`, await action('Load more').locator('.fui-Spinner__spinner').evaluate(node => node.offsetWidth), diameter);
      if (size === 'large') await primary.screenshot({ path: path.join(evidence, `${boundary}-load-more-loading.png`) });
    }
    await select('size').selectOption('medium');
    check('load-more spinner owns icon slot', await action('Load more').getByRole('progressbar', { name: 'Loading', exact: true }).count(), 1);
    check('load-more pending preserves rows', (await summary.textContent()).trim(), 'Loaded items 1–6 of 23. Loading requested results; current rows remain visible.');
    await primary.locator('[data-pagination-demo="complete"]').click();
    check('load-more completion appends once', (await summary.textContent()).trim(), 'Loaded items 1–9 of 23.');
    await checkbox('Complete requests manually').uncheck();
    await select('appearance').selectOption('subtle');
    await checkbox('Reduce total to three').check();
    check('incremental exhaustion', await action('Load more').isDisabled(), true);
    await select('mode').selectOption('known');
    check('one page actions disabled', await nav.locator('button').evaluateAll(nodes => nodes.every(node => node.disabled)), true);
    await select('filter').selectOption('missing');
    check('empty result summary', (await summary.textContent()).trim(), 'No matching orders.');
    check('empty page actions disabled', await nav.locator('button').evaluateAll(nodes => nodes.every(node => node.disabled)), true);
    await select('filter').selectOption('all');
    await checkbox('Reduce total to three').uncheck();
    await checkbox('Custom English labels').check();
    check('localized labels', await action('Next page').count(), 1);
    await checkbox('Hide navigation').check();
    check('hidden stays hidden', await nav.isVisible(), false);
    await primary.locator('[data-pagination-demo="inspect-ref"]').click();
    check('native hidden ref', (await primary.locator('[data-pagination-demo="ref"]').textContent()).trim(), 'Native ref: NAV; hidden=true; busy=absent');
    await checkbox('Hide navigation').uncheck();
    check('no accidental form submit', (await primary.locator('[data-pagination-demo="submissions"]').textContent()).trim(), 'Form submissions: 0');
    await primary.locator('[data-pagination-demo="submit"]').click();
    check('explicit form submit', (await primary.locator('[data-pagination-demo="submissions"]').textContent()).trim(), 'Form submissions: 1');
    const hostConfig = await page.evaluate(() => window.demoFixtureHost.config);
    const retainedNavigation = await nav.elementHandle();
    const primaryState = await summary.textContent();
    const secondaryState = await secondary.locator('[data-pagination-demo="summary"]').textContent();
    await page.evaluate(() => window.demoFixtureHost.update({ theme: 'dark' }));
    try {
      check('inherited dark token', await secondary.locator('nav').evaluate(node => getComputedStyle(node).getPropertyValue('--colorNeutralForeground1').trim()), teamsDarkTheme.colorNeutralForeground1);
      check('inherited RTL', await secondary.locator('nav').evaluate(node => getComputedStyle(node).direction), 'rtl');
      check('mirrored logical icon', await secondary.locator('nav button svg').first().evaluate(node => getComputedStyle(node).transform), 'matrix(-1, 0, 0, 1, 0, 0)');
      check('theme retains navigation DOM', await nav.evaluate((node, previous) => node === previous, retainedNavigation), true);
      check('theme retains primary result state', await summary.textContent(), primaryState);
      check('theme retains independent result state', await secondary.locator('[data-pagination-demo="summary"]').textContent(), secondaryState);
      check('theme retains form submissions', (await primary.locator('[data-pagination-demo="submissions"]').textContent()).trim(), 'Form submissions: 1');
      await page.screenshot({ path: path.join(evidence, `${boundary}-dark-rtl.png`), fullPage: true });
    } finally {
      await page.evaluate(config => window.demoFixtureHost.update(config), hostConfig);
      await retainedNavigation.dispose();
    }
    check('host configuration restored', await page.evaluate(() => window.demoFixtureHost.config), hostConfig);
    check('repository root uses classes', await nav.evaluate(node => node.hasAttribute('style')), false);
    await page.screenshot({ path: path.join(evidence, `${boundary}-desktop.png`), fullPage: true });
    await checkbox('320px navigation').check();
    await select('size').selectOption('small');
    await page.setViewportSize({ width: 320, height: 850 });
    await nav.scrollIntoViewIfNeeded();
    check('mobile fits available width', await nav.evaluate(node => node.getBoundingClientRect().width <= 288), true);
    check('mobile does not overflow document', await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
    check('mobile actions fit navigation', await nav.evaluate(node => {
      const area = node.getBoundingClientRect();
      return [...node.querySelectorAll('button')].every(button => {
        const box = button.getBoundingClientRect(); return box.left >= area.left && box.right <= area.right;
      });
    }), true);
    await action('Next page').click();
    check('mobile action works', (await summary.textContent()).trim(), 'Page 2: items 4–6 of 23.');
    await page.screenshot({ path: path.join(evidence, `${boundary}-mobile.png`), fullPage: true });
    const primarySummary = await summary.textContent();
    await secondary.locator('nav').getByRole('button', { name: 'Next page', exact: true }).click();
    await page.getByRole('button', { name: 'Independent RTL instance', exact: true }).click();
    check('collapsed independent navigation inaccessible', await secondary.locator('nav').evaluate(node => node.closest('.fui-AccordionPanel').hasAttribute('inert')), true);
    await page.getByRole('button', { name: 'Independent RTL instance', exact: true }).click();
    check('independent page retained after collapse', (await secondary.locator('[data-pagination-demo=summary]').textContent()).trim(), 'Page 2: items 6–10 of 23.');
    check('primary results isolated after secondary change', await summary.textContent(), primarySummary);
    check('console errors', errors, []);
    check('console warnings', warnings, []);
    return { boundary, url, checks, errors, warnings, browser: browser.version(), viewports: [{ width: 1440, height: 1100 }, { width: 320, height: 850 }] };
  } finally {
    await browser?.close();
    await new Promise(resolve => server.close(resolve));
  }
}

async function main() {
  fs.mkdirSync(evidence, { recursive: true });
  const reports = [];
  for (const boundary of ['source', 'compiled']) {
    const report = await verify(boundary);
    reports.push(report);
    console.log(`Pagination ${boundary}: ${report.checks.length} browser checks passed (${report.browser}).`);
  }
  fs.writeFileSync(path.join(evidence, 'browser-results.json'), JSON.stringify({ browserPlugin: 'not available; existing Playwright and Chrome used', reports }, undefined, 2));
}
main().catch(error => { console.error(error); process.exitCode = 1; });
