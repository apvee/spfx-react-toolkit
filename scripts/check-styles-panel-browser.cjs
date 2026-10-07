// Actual sample-panel check. Browser plugin unavailable: use explicitly supplied
// cached Playwright and installed Chrome; do not install or simulate a CSS engine.
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const { teamsLightTheme, teamsDarkTheme, teamsHighContrastTheme, webLightTheme } = require('@fluentui/react-theme');
const root = path.resolve(__dirname, '..');
const fixtureRoot = path.join(root, 'temp/styles-panel-browser');
const output = path.resolve(process.env.STYLES_PANEL_EVIDENCE || '.docs/maintenance/evidence/use-sx');
const label = process.env.STYLES_PANEL_LABEL || 'task-5-ui-browser';
const port = Number(process.env.STYLES_PANEL_PORT || 4320);
const url = `http://127.0.0.1:${port}`;
const observations = [], failures = [], consoleMessages = [], pageErrors = [];
let browser, server, page;
function check(name, actual, expected) {
  const pass = JSON.stringify(actual) === JSON.stringify(expected);
  observations.push({ name, actual, expected, pass });
  if (!pass) failures.push(name);
}
const selector = id => `[data-sx-demo="${id}"]`;
const node = id => page.locator(selector(id));
const computed = (id, prop) => node(id).evaluate((element, property) => getComputedStyle(element)[property], prop);
async function select(id, value) { await node(id).selectOption(value); }
async function fill(id, value) { await node(id).fill(String(value)); }
async function settle() { await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))); }
async function token(name) { return node('provider').evaluate((element, value) => getComputedStyle(element).getPropertyValue(`--${value}`).trim(), name); }
const rgb = hex => /^#[\da-f]{6}$/i.test(hex) ? `rgb(${[1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16)).join(', ')})` : hex;
async function themeChecks(choice, theme) {
  await select('theme', choice); await settle();
  check(`${choice}:real-theme-token`, await token('colorBrandBackground'), theme.colorBrandBackground);
  await select('preset', 'canvas');
  check(`${choice}:canvas-background`, await computed('preset-preview', 'backgroundColor'), rgb(theme.colorNeutralBackground1));
  check(`${choice}:canvas-color`, await computed('preset-preview', 'color'), rgb(theme.colorNeutralForeground1));
}
(async () => {
  fs.mkdirSync(output, { recursive: true });
  server = http.createServer((request, response) => {
    const pathname = new URL(request.url, url).pathname;
    if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
    const file = path.resolve(fixtureRoot, pathname === '/' ? 'index.html' : '.' + pathname);
    if (!file.startsWith(fixtureRoot + path.sep)) { response.writeHead(403).end(); return; }
    fs.readFile(file, (error, data) => {
      if (error) { response.writeHead(404).end(); return; }
      response.setHeader('Content-Type', file.endsWith('.html') ? 'text/html; charset=utf-8' : 'application/javascript'); response.end(data);
    });
  });
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(port, '127.0.0.1', resolve); });
  browser = await chromium.launch({ executablePath: process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', headless: true, ignoreDefaultArgs: ['--hide-scrollbars'] });
  page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  page.on('console', message => { if (message.type() === 'error' || message.type() === 'warning') consoleMessages.push({ type: message.type(), text: message.text() }); });
  page.on('pageerror', error => pageErrors.push(error.message));
  await page.goto(url); await node('layout-preview').waitFor();
  check('page identity', await page.title(), 'StylesPanel local browser fixture');
  check('local origin', new URL(page.url()).origin, url);
  check('meaningful actual panel', await page.locator('h3').filter({ hasText: 'Layout and roles' }).count(), 1);
  check('framework overlay', await page.locator('webpack-dev-server-client-overlay,#webpack-dev-server-client-overlay,nextjs-portal').count(), 0);
  check('sample-owned probes have no inline styles', await page.locator('[data-sx-demo][style]').count(), 0);
  check('fixture discloses SPFx boundary', (await page.locator('body').textContent()).includes('not authenticated SharePoint validation'), true);
  await page.screenshot({ path: path.join(output, `${label}-desktop.png`), fullPage: true });
  await fill('layout-width', 240.5); await fill('layout-gap', 24.5);
  check('fractional width accepted', await computed('layout-preview', 'width'), '240.5px');
  check('fractional gap accepted', await computed('layout-preview', 'columnGap'), '24.5px');
  await fill('layout-width', 360); await fill('layout-gap', 24); await fill('layout-columns', 3);
  check('width control', await computed('layout-preview', 'width'), '360px');
  check('gap row control', await computed('layout-preview', 'rowGap'), '24px');
  check('gap column control', await computed('layout-preview', 'columnGap'), '24px');
  check('grid columns control', (await computed('layout-preview', 'gridTemplateColumns')).split(' ').length, 3);
  const widthClasses = await node('layout-preview').getAttribute('class');
  await fill('layout-width', 240); check('width control resets', await computed('layout-preview', 'width'), '240px');
  await fill('layout-width', 360); check('repeated width reuses classes', await node('layout-preview').getAttribute('class'), widthClasses);
  await select('foreground', 'subtle'); check('foreground subtle', await computed('role-preview', 'color'), rgb(webLightTheme.colorNeutralForeground2));
  await select('background', 'subtle'); check('subtle background role', await computed('role-preview', 'backgroundColor'), 'rgba(0, 0, 0, 0)');
  await select('background', 'alternative'); check('alternative background role', await computed('role-preview', 'backgroundColor'), rgb(webLightTheme.colorNeutralBackground2));
  await select('preset', 'brand'); const presetClasses = await node('preset-preview').getAttribute('class');
  await select('theme', 'dark'); check('theme retains descriptor classes', await node('preset-preview').getAttribute('class'), presetClasses);
  check('dark preset live variables', await computed('preset-preview', 'backgroundColor'), rgb(teamsDarkTheme.colorBrandBackground));
  await themeChecks('host', webLightTheme); await themeChecks('default', teamsLightTheme);
  await themeChecks('dark', teamsDarkTheme); await themeChecks('contrast', teamsHighContrastTheme);
  await select('preset', 'inverted'); check('high contrast inverted unavailable', await node('preset-preview').getAttribute('data-unsupported'), 'true');
  check('high contrast visible notice', (await node('preset-notice').textContent()).includes('unsupported'), true);
  check('high contrast readable canvas fallback', await computed('preset-preview', 'backgroundColor'), rgb(teamsHighContrastTheme.colorNeutralBackground1));
  await select('preset', 'disabled'); check('disabled preset semantic presentation', await node('preset-preview').getAttribute('aria-disabled'), 'true');
  await select('theme', 'default'); await select('direction', 'rtl');
  check('active provider RTL logical padding', await computed('direction-preview', 'paddingRight'), '32px');
  check('active provider RTL opposite padding', await computed('direction-preview', 'paddingLeft'), '12px');
  await select('direction', 'ltr'); check('active provider LTR logical padding', await computed('direction-preview', 'paddingLeft'), '32px');
  await select('responsive-threshold', 'medium');
  check('nearest independent small region', await computed('query-a-preview', 'width'), '180px');
  check('nearest independent large region', await computed('query-b-preview', 'width'), '300px');
  await fill('query-a-width', 720); await fill('query-b-width', 420); await settle();
  check('region A independently grows', await computed('query-a-preview', 'width'), '300px');
  check('region B independently shrinks', await computed('query-b-preview', 'width'), '180px');
  for (const [threshold, limit] of [['small', 480], ['medium', 640], ['large', 1024]]) {
    await select('responsive-threshold', threshold); await fill('query-a-width', limit - 1); await settle();
    check(`${threshold}:query below`, await computed('query-a-preview', 'width'), '180px');
    await fill('query-a-width', limit); await settle();
    check(`${threshold}:query inclusive`, await computed('query-a-preview', 'width'), '300px');
    await select('viewport-threshold', threshold); await page.setViewportSize({ width: limit - 1, height: 1000 }); await settle();
    check(`${threshold}:viewport below`, await computed('viewport-preview', 'width'), '180px');
    await page.setViewportSize({ width: limit, height: 1000 }); await settle();
    check(`${threshold}:viewport inclusive`, await computed('viewport-preview', 'width'), '300px');
  }
  await page.setViewportSize({ width: 1440, height: 1000 }); await fill('query-a-width', 420); await fill('query-b-width', 720);
  await node('selected').check(); await node('selected').focus(); await page.keyboard.press('Tab');
  check('keyboard reaches native action', await node('state-preview').evaluate(element => document.activeElement === element), true);
  check('real focus-visible state', await node('state-preview').evaluate(element => element.matches(':focus-visible')), true);
  check('focus recipe border', await computed('state-preview', 'borderTopColor'), rgb(teamsLightTheme.colorBrandStroke1));
  check('native visible outline', await node('state-preview').evaluate(element => { const s = getComputedStyle(element); return parseFloat(s.outlineWidth) > 0 && s.outlineStyle !== 'none'; }), true);
  await node('state-preview').hover({ position: { x: 5, y: 5 } }); await page.mouse.down();
  check('simultaneous native states', await node('state-preview').evaluate(element => [':hover', ':active', ':focus-visible'].map(state => element.matches(state))), [true, true, true]);
  check('active wins hover and selected', await computed('state-preview', 'backgroundColor'), rgb(teamsLightTheme.colorBrandBackground));
  await page.mouse.up(); await node('disabled').check();
  check('native disabled', await node('state-preview').isDisabled(), true);
  const before = await page.getByText('Action invocations:', { exact: false }).textContent();
  await node('state-preview').hover({ position: { x: 5, y: 5 } });
  check('disabled hover drops enabled recipe', await computed('state-preview', 'backgroundColor'), rgb(teamsLightTheme.colorNeutralBackgroundDisabled));
  await node('state-preview').evaluate(element => element.click());
  check('disabled native action inert', await page.getByText('Action invocations:', { exact: false }).textContent(), before);
  await node('disabled').uncheck(); await node('selected').uncheck();
  check('scroll recipe overflow', await computed('scroll-preview', 'overflowY'), 'auto');
  check('scroll recipe width', await computed('scroll-preview', 'scrollbarWidth'), 'thin');
  check('scroll region genuinely scrolls', await node('scroll-preview').evaluate(element => { element.scrollTop = 40; return element.scrollTop; }), 40);
  await page.emulateMedia({ forcedColors: 'active' });
  check('forced colors real media active', await page.evaluate(() => matchMedia('(forced-colors: active)').matches), true);
  check('forced colors system scrollbar', await computed('scroll-preview', 'scrollbarColor'), 'auto');
  check('forced color adjust retained', await computed('scroll-preview', 'forcedColorAdjust'), 'auto');
  await page.emulateMedia({ forcedColors: 'none' });
  const catalog = await node('catalog-family').locator('option').evaluateAll(options => options.map(option => option.value));
  check('catalog has all base namespace choices', catalog.length, 43);
  let exercised = 0;
  for (const family of catalog) {
    await select('catalog-family', family);
    const members = await node('catalog-member').locator('option').evaluateAll(options => options.map(option => option.value));
    let previous;
    for (const member of members) {
      await select('catalog-member', member);
      const className = await node('catalog-preview').getAttribute('class');
      check(`catalog ${family}.${member} real class`, Boolean(className && className.includes('f')), true);
      if (previous) check(`catalog ${family}.${member} changes applied descriptor`, className !== previous, true);
      previous = className; exercised++;
    }
  }
  check('catalog member count', exercised, 220);
  await select('catalog-family', 'presets'); await select('catalog-member', 'inverted'); await select('theme', 'contrast');
  check('catalog inverted high contrast unavailable', await node('catalog-preview').getAttribute('data-unsupported'), 'true');
  await page.screenshot({ path: path.join(output, `${label}-high-contrast.png`), fullPage: true });
  await select('theme', 'dark'); await select('preset', 'canvas'); await select('catalog-member', 'canvas');
  await page.screenshot({ path: path.join(output, `${label}-dark.png`), fullPage: true });
  await page.setViewportSize({ width: 390, height: 844 }); await settle();
  await page.screenshot({ path: path.join(output, `${label}-mobile.png`), fullPage: true });
  check('mobile controls stay usable', await node('catalog-family').isVisible(), true);
  check('mobile no page-wide overflow', await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true);
  check('console health', consoleMessages, []); check('page errors', pageErrors, []);
})().catch(error => { failures.push(error.stack || error.message); }).finally(async () => {
  const version = browser ? browser.version() : 'not launched';
  if (browser) await browser.close();
  if (server) await new Promise(resolve => server.close(resolve));
  fs.mkdirSync(output, { recursive: true });
  fs.writeFileSync(path.join(output, `${label}.json`), JSON.stringify({ url, browser: version, boundary: 'SPFx Teams/theme hook context only; actual compiled library, sample panel, React17, Griffel, FluentProvider', observations, failures, consoleMessages, pageErrors }, null, 2));
  console.log(JSON.stringify({ browser: version, checks: observations.length, failures, consoleMessages, pageErrors }, null, 2));
  if (failures.length) process.exitCode = 1;
});
