// Actual sample-panel check. Browser plugin unavailable: use explicitly supplied
// cached Playwright and installed Chrome; do not install or simulate a CSS engine.
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const { teamsLightTheme, teamsDarkTheme, teamsHighContrastTheme, webLightTheme } = require('@fluentui/react-theme');
const root = path.resolve(__dirname, '..');
const fixtureRoot = path.join(root, 'temp/styles-panel-browser');
const output = path.resolve(process.env.STYLES_PANEL_EVIDENCE || (process.env.STYLES_PANEL_STAGE_ONLY === '1' ? '.docs/maintenance/evidence/stage' : process.env.STYLES_PANEL_SHELL_ONLY === '1' ? '.docs/maintenance/evidence/web-part-shell' : '.docs/maintenance/evidence/use-sx'));
const label = process.env.STYLES_PANEL_LABEL || (process.env.STYLES_PANEL_STAGE_ONLY === '1' ? 'stage-compiled' : process.env.STYLES_PANEL_SHELL_ONLY === '1' ? 'shell-compiled' : 'task-5-ui-browser');
const port = Number(process.env.STYLES_PANEL_PORT || 4320);
const url = `http://127.0.0.1:${port}`;
const observations = [], failures = [], consoleMessages = [], pageErrors = [];
const thumbInteractionCaptures = [];
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
async function thumbInteractions(axis) {
  const target = node('scroll-preview');
  await target.scrollIntoViewIfNeeded();
  await target.evaluate(element => { element.scrollTop = 0; element.scrollLeft = 0; });
  const geometry = await target.evaluate(element => {
    const rect = element.getBoundingClientRect(), css = getComputedStyle(element);
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height,
      left: parseFloat(css.borderLeftWidth), right: parseFloat(css.borderRightWidth),
      top: parseFloat(css.borderTopWidth), bottom: parseFloat(css.borderBottomWidth) };
  });
  const point = axis === 'vertical'
    ? { x: geometry.x + geometry.width - geometry.right - 3, y: geometry.y + geometry.top + 10 }
    : { x: geometry.x + geometry.left + 10, y: geometry.y + geometry.height - geometry.bottom - 3 };
  let pointer = { x: 0, y: 0 };
  async function capture(phase, colorToken) {
    await settle();
    const filename = `${label}-scroll-thumb-${axis}-${phase}.png`;
    await target.screenshot({ path: path.join(output, filename) });
    thumbInteractionCaptures.push({ axis, phase, filename, expectedColor: rgb(await token(colorToken)),
      pointer, geometry, paintVerification: 'Requires screenshot/pixel review; computed pseudo style does not expose native thumb state' });
  }
  await page.mouse.move(0, 0);
  await capture('base', 'colorNeutralStrokeAccessible');
  pointer = point;
  await page.mouse.move(point.x, point.y);
  await capture('hover', 'colorNeutralStrokeAccessibleHover');
  await page.mouse.down();
  try {
    await capture('pressed', 'colorNeutralStrokeAccessiblePressed');
    pointer = { x: point.x + (axis === 'horizontal' ? 30 : 0), y: point.y + (axis === 'vertical' ? 15 : 0) };
    await page.mouse.move(pointer.x, pointer.y, { steps: 4 });
    await capture('drag', 'colorNeutralStrokeAccessiblePressed');
    check(`actual ${axis} thumb drag scrolls`, await target.evaluate((element, direction) =>
      direction === 'vertical' ? element.scrollTop > 0 : element.scrollLeft > 0, axis), true);
  } finally { await page.mouse.up(); }
  await capture('released', 'colorNeutralStrokeAccessibleHover');
  pointer = { x: 0, y: 0 };
  await page.mouse.move(0, 0);
  await capture('leave', 'colorNeutralStrokeAccessible');
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
  if (process.env.STYLES_PANEL_STAGE_ONLY === '1') {
    await require('./stage-browser.helpers.cjs').checkStageBrowser(page, { url, output, label, boundary: 'compiled', check, consoleMessages });
    check('Stage console health', consoleMessages, []);
    check('Stage page errors', pageErrors, []);
    return;
  }
  if (process.env.STYLES_PANEL_SHELL_ONLY === '1') {
    await require('./web-part-shell-browser.helpers.cjs').checkWebPartShellBrowser(page, { url, output, label, boundary: 'compiled', check, consoleMessages });
    check('Shell console health', consoleMessages, []);
    check('Shell page errors', pageErrors, []);
    return;
  }
  if (process.env.DATA_GRID_CHART_PROBE === '1') {
    await require('./data-grid-browser.helpers.cjs').exerciseDataGridChartProbe({ page, url, output, label, boundary: 'compiled', check, consoleMessages, pageErrors });
    return;
  }
  if (process.env.STYLES_PANEL_DATA_GRID_ONLY === '1') {
    await require('./data-grid-browser.helpers.cjs').exerciseDataGrid({ page, url, output, label, boundary: 'compiled', check, consoleMessages });
    check('DataGrid page errors', pageErrors, []);
    return;
  }
  if (process.env.STYLES_PANEL_FORM_ONLY === '1') {
    await require('./form-browser.helpers.cjs').exerciseForm({ page, url, output, label, boundary: 'compiled', check, consoleMessages });
    check('Form page errors', pageErrors, []);
    return;
  }
  await page.goto(url); await node('layout-preview').waitFor();
  check('page identity', await page.title(), 'StylesPanel local browser fixture');
  check('local origin', new URL(page.url()).origin, url);
  check('meaningful actual panel', await page.locator('h3').filter({ hasText: 'Layout and roles' }).count(), 1);
  check('framework overlay', await page.locator('webpack-dev-server-client-overlay,#webpack-dev-server-client-overlay,nextjs-portal').count(), 0);
  check('sample-owned probes have no inline styles', await page.locator('[data-sx-demo][style]').count(), 0);
  check('fixture discloses SPFx boundary', (await page.locator('body').textContent()).includes('not authenticated SharePoint validation'), true);
  await page.screenshot({ path: path.join(output, `${label}-desktop.png`), fullPage: true });
  if (process.env.STYLES_PANEL_GROUPS_ONLY !== '1') {
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
  const vendorScrollbars = await page.evaluate(() => CSS.supports('selector(::-webkit-scrollbar)'));
  check('scroll recipe width', await computed('scroll-preview', 'scrollbarWidth'), vendorScrollbars ? 'auto' : 'thin');
  if (vendorScrollbars) check('scroll recipe equally sized axes', await node('scroll-preview').evaluate(element => {
    const css = getComputedStyle(element, '::-webkit-scrollbar');
    return { width: css.width, height: css.height };
  }), { width: '6px', height: '6px' });
  check('scroll region genuinely scrolls', await node('scroll-preview').evaluate(element => { element.scrollTop = 40; return element.scrollTop; }), 40);
  check('scroll region scrolls horizontally', await node('scroll-preview').evaluate(element => { element.scrollLeft = 40; return element.scrollLeft; }), 40);
  if (vendorScrollbars) {
    check('actual thumb selectors supported', await page.evaluate(() =>
      ['::-webkit-scrollbar-thumb:hover', '::-webkit-scrollbar-thumb:active', '::-webkit-scrollbar-thumb:hover:active']
        .every(selector => CSS.supports(`selector(${selector})`))), true);
    for (const axis of ['vertical', 'horizontal']) await thumbInteractions(axis);
  }
  await page.emulateMedia({ forcedColors: 'active' });
  check('forced colors real media active', await page.evaluate(() => matchMedia('(forced-colors: active)').matches), true);
  check('forced colors system scrollbar', await computed('scroll-preview', 'scrollbarColor'), 'auto');
  check('forced colors native thin width', await computed('scroll-preview', 'scrollbarWidth'), 'thin');
  if (vendorScrollbars) check('forced colors vendor sizing inactive', await node('scroll-preview').evaluate(element => {
    const css = getComputedStyle(element, '::-webkit-scrollbar');
    return { width: css.width, height: css.height };
  }), { width: 'auto', height: 'auto' });
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
  // Actual FormattedText sample consumes the compiled library through the existing boundary.
  const ft = id => page.locator(`[data-formatted-text-demo="${id}"]`);
  await ft('text').fill('Read **updates** in [docs](#local). ' + 'Long text '.repeat(40));
  await ft('variant').selectOption('title2');
  check('FormattedText variant keeps default span', await ft('preview').evaluate(e => e.tagName), 'SPAN');
  await ft('as').selectOption('h3');
  check('FormattedText explicit semantics', await ft('preview').evaluate(e => e.tagName), 'H3');
  await ft('inspect-ref').click();
  check('FormattedText native ref observation', (await ft('observations').textContent()).includes('Ref element: h3.'), true);
  await ft('markdown').uncheck();
  check('FormattedText Markdown disabled', await ft('preview').locator('strong').count(), 0);
  await ft('markdown').check();
  check('FormattedText Markdown enabled', await ft('preview').locator('strong').count(), 1);
  await ft('maxLines').fill('2');
  check('FormattedText sample clamp', await ft('preview').evaluate(e => getComputedStyle(e).webkitLineClamp), '2');
  await ft('subtle').check(); await ft('highlight').check();
  check('FormattedText sample highlight', await ft('preview').locator('mark').count(), 1);
  await ft('preview').locator('strong').click();
  check('FormattedText sample native click', (await ft('observations').textContent()).includes('Preview clicks: 1.'), true);
  await ft('rtl').check(); await ft('theme').check();
  check('FormattedText sample RTL', await ft('preview').evaluate(e => getComputedStyle(e).direction), 'rtl');
  check('FormattedText preserves authored JSX', await ft('mixed-jsx').locator('button').count(), 1);
  await ft('mixed-jsx').locator('button').click();
  check('FormattedText authored JSX click', (await ft('observations').textContent()).includes('Preview clicks: 2.'), true);
  await ft('preview').screenshot({ path: path.join(output, `${label}-formatted-text-preview.png`) });
  const rt = id => page.locator(`[data-relative-time-demo="${id}"]`);
  await rt('dateTime').fill('2026-01-31T12:00:00Z');
  await rt('referenceTime').fill('2026-02-28T12:00:00Z');
  await rt('locale').selectOption('en-US');
  await rt('mode').selectOption('fixed');
  await rt('showTooltip').uncheck();
  check('RelativeTime sample native tag', await rt('preview').evaluate(e => e.tagName), 'TIME');
  check('RelativeTime sample calendar month', await rt('preview').textContent(), '1 month ago');
  await rt('inspect-ref').click();
  check('RelativeTime sample native ref', (await rt('observations').textContent()).includes('Ref element: time.'), true);
  await rt('preview').click();
  check('RelativeTime sample native click', (await rt('observations').textContent()).includes('Preview clicks: 1.'), true);
  await rt('locale').selectOption('it-IT');
  check('RelativeTime sample locale', await rt('preview').textContent(), await page.evaluate(() => new Intl.RelativeTimeFormat('it-IT', { numeric: 'always' }).format(-1, 'month')));
  await rt('dateTime').fill('invalid');
  check('RelativeTime sample invalid fallback', await rt('preview').count(), 0);
  check('RelativeTime sample visible fallback', (await page.locator('body').textContent()).includes('Date unavailable'), true);
  await rt('dateTime').fill('2026-01-31T12:00:00Z');
  await rt('mode').selectOption('live');
  check('RelativeTime sample live resumes', await rt('preview').count(), 1);
  await rt('mode').selectOption('frozen');
  check('RelativeTime sample frozen usable', await rt('preview').count(), 1);
  await rt('mode').selectOption('fixed');
  await rt('preview').screenshot({ path: path.join(output, `${label}-relative-time-preview.png`) });
  const sk = id => page.locator(`[data-auto-skeleton-demo="${id}"]`);
  const skPreview = sk('preview');
  await sk('focus-target').fill('retained sample note');
  await sk('inspect-ref').click(); check('AutoSkeleton sample native ref',(await sk('observations').textContent()).includes('Root: div,'),true);
  await sk('focus-load').click(); await settle();
  check('AutoSkeleton sample busy',await skPreview.getAttribute('aria-busy'),'true');
  check('AutoSkeleton sample overlay',await skPreview.locator('[data-auto-skeleton-overlay]').count(),1);
  check('AutoSkeleton sample exclusive shapes',await skPreview.locator('[data-auto-skeleton-shape]').evaluateAll(items=>items.every(e=>['circle','rectangle'].includes(e.getAttribute('data-auto-skeleton-shape')))),true);
  check('AutoSkeleton sample rectangle dimensions',await skPreview.locator('[data-auto-skeleton-shape]').evaluateAll(items=>items.some(e=>{const r=e.getBoundingClientRect();return r.width===160 && r.height===24;})),true);
  check('AutoSkeleton sample inert',await skPreview.locator('[data-auto-skeleton-content]').evaluate(e=>e.inert),true);
  check('AutoSkeleton sample focus released',await page.evaluate(()=>!document.activeElement.matches('[data-auto-skeleton-demo="focus-target"]')),true);
  const count=await skPreview.locator('[data-auto-skeleton-shape]').count(); await sk('resize').check(); await settle();
  check('AutoSkeleton sample loading resize',await skPreview.locator('[data-auto-skeleton-shape]').count()>count,true);
  await sk('animation').selectOption('pulse'); await settle();
  await sk('direction').check(); await settle();
  check('AutoSkeleton sample RTL geometry',await skPreview.evaluate(root=>{const rect=root.querySelector('[data-auto-skeleton-demo="rectangle"]').getBoundingClientRect();return [...root.querySelectorAll('[data-auto-skeleton-shape]')].some(e=>{const r=e.getBoundingClientRect();return r.width===160&&r.height===24&&Math.abs(r.left-rect.left)<=1&&Math.abs(r.top-rect.top)<=1;});}),true);
  await sk('theme').check(); await settle();
  await skPreview.screenshot({path:path.join(output,`${label}-auto-skeleton-preview.png`)});
  await sk('loading').uncheck(); await settle();
  check('AutoSkeleton sample state retained',await sk('focus-target').inputValue(),'retained sample note');
  check('AutoSkeleton sample mounted once',(await sk('observations').textContent()).includes('Content mounts: 1.'),true);
  await sk('loading').check(); await sk('empty').check(); await settle();
  check('AutoSkeleton sample empty clears',await skPreview.locator('[data-auto-skeleton-overlay]').count(),0);
  check('AutoSkeleton sample empty stays busy',await skPreview.getAttribute('aria-busy'),'true');
  const sm = id => page.locator(`[data-state-message-demo="${id}"]`);
  const smRoot=sm('preview');
  check('StateMessage sample h3',await smRoot.locator('h3').textContent(),'Your section state');
  check('StateMessage sample status',await smRoot.locator('[data-state-message-text]').getAttribute('role'),'status');
  await sm('inspect').click();check('StateMessage sample root ref',(await sm('observations').textContent()).includes('Root: div, heading=h3'),true);
  await smRoot.getByRole('button',{name:'Try again'}).click();check('StateMessage sample action',(await sm('observations').textContent()).includes('Actions: 1.'),true);
  check('StateMessage sample disabled',await smRoot.getByRole('button',{name:'Unavailable'}).isDisabled(),true);
  for(const preset of ['generic','empty','error','success','welcome']){
    await sm('preset').selectOption(preset);await settle();
    check(`StateMessage sample ${preset} artwork`,await smRoot.locator('svg').count(),1);
  }
  await sm('theme').check();await sm('direction').check();await sm('resize').check();await sm('long').check();await settle();
  check('StateMessage sample narrow no overflow',await smRoot.evaluate(e=>e.scrollWidth<=e.clientWidth),true);
  check('StateMessage sample rtl',await smRoot.evaluate(e=>getComputedStyle(e).direction),'rtl');
  await sm('role').check();check('StateMessage sample explicit alert',await smRoot.locator('[data-state-message-text]').getAttribute('role'),'alert');
  await sm('heading').selectOption('2');check('StateMessage sample semantic heading',await smRoot.locator('h2').count(),1);
  await sm('visual').selectOption('hidden');check('StateMessage sample hide',await smRoot.locator('[data-state-message-visual]').count(),0);
  await sm('visual').selectOption('custom');check('StateMessage sample custom',await smRoot.locator('[role="img"]').getAttribute('aria-label'),'Custom visual');
  await sm('visual').selectOption('preset');await smRoot.screenshot({path:path.join(output,`${label}-state-message-preview.png`)});
  }
  await page.setViewportSize({ width: 1440, height: 1000 }); await settle();
  const groups = id => page.locator(`[data-groups-demo="${id}"]`);
  const settings = groups('settings'), accordion = groups('accordion');
  const profile = accordion.locator('[data-accordion-group-item]').first();
  const advanced = accordion.locator('[data-accordion-group-item]').nth(1);
  const profileTrigger = profile.locator('button').first();
  const advancedTrigger = advanced.locator('button').first();
  for (const size of [500, 639, 640]) {
    await groups('resize').selectOption(String(size)); await settle();
    check(`Groups sample ${size} real container width`, await settings.evaluate(e => e.clientWidth), size);
    check(`Groups sample ${size} fluid rows`, await settings.locator('[data-settings-group-row]').first().evaluate(e => getComputedStyle(e).flexWrap), 'wrap');
    check(`Groups sample ${size} compact switch alongside`, await settings.locator('[data-settings-group-item]').nth(1).evaluate(e => {
      const header = e.querySelector('[data-settings-group-leading]').getBoundingClientRect(), controls = e.querySelector('[data-settings-group-controls]').getBoundingClientRect();
      return Math.abs((header.top + header.bottom) - (controls.top + controls.bottom)) < 2;
    }), true);
    if (size === 500) await settings.screenshot({ path: path.join(output, `${label}-groups-500.png`) });
  }
  check('Groups sample single caller control per settings row', await settings.locator('[data-settings-group-row]').evaluateAll(nodes => nodes.length === 2 && nodes.every(e => e.children.length === 2 && e.querySelectorAll('input').length === 1 && !e.querySelector('.fui-Field, button'))), true);
  check('Groups sample zero text gaps and margins', await page.locator('[data-groups-demo="frame"] [data-group-text]').evaluateAll(nodes => nodes.every(e => getComputedStyle(e).rowGap === '0px' && [...e.children].every(child => { const css = getComputedStyle(child); return css.marginTop === '0px' && css.marginBottom === '0px'; }))), true);
  check('Groups sample divider derives total Card spacing', await page.locator('[data-groups-demo="settings"] [data-settings-group-item] + [data-settings-group-item], [data-groups-demo="accordion"] [data-accordion-group-item] + [data-accordion-group-item]').evaluateAll(nodes => nodes.length === 2 && nodes.every(e => {
    const css = getComputedStyle(e), previous = e.previousElementSibling, padding = parseFloat(getComputedStyle(e.closest('.fui-Card')).paddingTop);
    const geometry = e.getBoundingClientRect().top - previous.getBoundingClientRect().bottom + parseFloat(css.paddingTop) + parseFloat(css.borderTopWidth);
    return padding === 12 && css.borderTopWidth === '1px' && parseFloat(css.paddingTop) === padding / 2 && parseFloat(css.marginTop) === padding / 2 && geometry === padding + 1;
  })), true);
  check('Groups sample compact Settings row density', await settings.locator('[data-settings-group-row]').evaluateAll(nodes => nodes.every(e => e.getBoundingClientRect().height === 44)), true);
  check('Groups sample compact Accordion row density', await accordion.locator('.fui-AccordionHeader button').evaluateAll(nodes => nodes.every(e => e.getBoundingClientRect().height === 44)), true);
  check('Groups sample single Card inset', await page.locator('[data-groups-demo="frame"] .fui-Card').evaluateAll(nodes => nodes.every(e => { const css = getComputedStyle(e); return css.paddingLeft === '12px' && css.paddingRight === '12px' && css.rowGap === '0px'; })), true);
  check('Groups sample header Card gap 12px', await page.locator('[data-groups-demo="settings"], [data-groups-demo="accordion"]').evaluateAll(nodes => nodes.every(e => getComputedStyle(e).rowGap === '12px')), true);
  check('Groups sample Settings divider row-box symmetry', await settings.locator('[data-settings-group-item]').nth(1).evaluate(e => {
    const before = e.previousElementSibling.querySelector('[data-settings-group-row]').getBoundingClientRect(), after = e.querySelector('[data-settings-group-row]').getBoundingClientRect();
    const divider = e.getBoundingClientRect(), stroke = parseFloat(getComputedStyle(e).borderTopWidth);
    return divider.top - before.bottom === 6 && after.top - divider.top - stroke === 6;
  }), true);
  check('Groups sample Accordion divider body-row symmetry', await advanced.evaluate(e => {
    const before = e.previousElementSibling.getBoundingClientRect(), after = e.querySelector('.fui-AccordionHeader button').getBoundingClientRect();
    const divider = e.getBoundingClientRect(), stroke = parseFloat(getComputedStyle(e).borderTopWidth);
    return divider.top - before.bottom === 6 && after.top - divider.top - stroke === 6 && after.height === 44;
  }), true);
  check('Groups sample zero Accordion lateral inset', await accordion.evaluate(e => [...e.querySelectorAll('.fui-AccordionHeader button')].every(button => { const css = getComputedStyle(button); return css.paddingLeft === '0px' && css.paddingRight === '0px'; }) && [...e.querySelectorAll('.fui-AccordionPanel')].every(panel => { const css = getComputedStyle(panel); return css.marginLeft === '0px' && css.marginRight === '0px'; })), true);
  const iconDiagnostics = await page.locator('[data-groups-demo="frame"] [data-icon-name]').evaluateAll(nodes => nodes.map(e => {
    const dimensions = node => { const rect = node.getBoundingClientRect(), css = getComputedStyle(node); return { width: rect.width, height: rect.height, fontSize: css.fontSize, lineHeight: css.lineHeight, display: css.display }; };
    return { name: e.getAttribute('data-icon-name'), icon: dimensions(e), slot: dimensions(e.parentElement), svg: e.querySelector('svg') ? dimensions(e.querySelector('svg')) : undefined };
  }));
  observations.push({ name: 'Groups sample icon geometry diagnostic', actual: iconDiagnostics, expected: 'Diagnostic only; assertions below verify dimensions' });
  check('Groups sample library header icon slot and font 24px', await page.locator('[data-groups-demo="frame"] [data-group-icon]').evaluateAll(nodes => nodes.length === 2 && nodes.every(e => {
    const rect = e.getBoundingClientRect(), icon = e.querySelector('[data-icon-name]'), painted = icon.querySelector('svg').getBoundingClientRect();
    return rect.width === 24 && rect.height === 24 && getComputedStyle(icon).fontSize === '24px' && painted.width === 24 && painted.height === 24;
  })), true);
  check('Groups sample library item icon slot and font 20px', await page.locator('[data-groups-demo="frame"] [data-group-item-icon]').evaluateAll(nodes => nodes.length === 4 && nodes.every(e => {
    const rect = e.getBoundingClientRect(), icon = e.querySelector('[data-icon-name]'), painted = icon.querySelector('svg').getBoundingClientRect();
    return rect.width === 20 && rect.height === 20 && getComputedStyle(icon).fontSize === '20px' && painted.width === 20 && painted.height === 20;
  })), true);
  await groups('email').fill('edited@example.com');
  await groups('note').fill('Sample retained edit');
  await profileTrigger.click();
  check('Groups sample closes', await profileTrigger.getAttribute('aria-expanded'), 'false');
  check('Groups sample retained inert', await groups('note').evaluate(e => Boolean(e.closest('[inert]'))), true);
  await profileTrigger.focus(); await page.keyboard.press('Tab');
  check('Groups sample closed keyboard exclusion', await advancedTrigger.evaluate(e => e === document.activeElement), true);
  await profileTrigger.click();
  check('Groups sample input persists', await groups('note').inputValue(), 'Sample retained edit');
  await groups('close-all').focus(); await page.keyboard.press('Enter');
  check('Groups sample action closes controlled panel', await profileTrigger.getAttribute('aria-expanded'), 'false');
  check('Groups sample close action keeps visible focus', await groups('close-all').evaluate(e => e === document.activeElement && !e.closest('[inert]')), true);
  await profileTrigger.click();
  check('Groups sample action closure retains state', await groups('note').inputValue(), 'Sample retained edit');
  await groups('multiple').check();
  await advancedTrigger.click();
  check('Groups sample multiple profile stays open', await profileTrigger.getAttribute('aria-expanded'), 'true');
  check('Groups sample multiple advanced opens', await advancedTrigger.getAttribute('aria-expanded'), 'true');
  await groups('unmount').check();
  await groups('close-all').focus(); await page.keyboard.press('Enter');
  check('Groups sample unmount action keeps visible focus', await groups('close-all').evaluate(e => e === document.activeElement && !e.closest('[inert]')), true);
  await page.waitForFunction(() => !document.querySelector('[data-groups-demo="note"]'));
  check('Groups sample unmounted optional', await groups('note').count(), 0);
  await profileTrigger.click();
  check('Groups sample optional reset', await groups('note').inputValue(), 'Retained note');
  await groups('inspect').click();
  check('Groups sample native refs', (await groups('observations').textContent()).includes('settings=div, accordion=div'), true);
  await groups('disabled').check();
  check('Groups sample disabled', await profileTrigger.isDisabled(), true);
  await groups('disabled').uncheck();
  await groups('hidden').check();
  check('Groups sample native hidden', await settings.isVisible(), false);
  await groups('hidden').uncheck();
  await groups('theme').check(); await groups('direction').check(); await settle();
  check('Groups sample RTL', await settings.evaluate(e => getComputedStyle(e).direction), 'rtl');
  check('Groups sample theme token', await settings.evaluate(e => getComputedStyle(e).getPropertyValue('--colorNeutralForeground1').trim()), '#ffffff');
  check('Groups sample caller settings preserved', await groups('email').inputValue(), 'edited@example.com');
  await groups('theme').uncheck(); await groups('direction').uncheck(); await settle();
  await groups('frame').screenshot({ path: path.join(output, `${label}-groups-light.png`) });
  await groups('theme').check(); await groups('direction').check(); await settle();
  await groups('frame').screenshot({ path: path.join(output, `${label}-groups-preview.png`) });
  await groups('resize').selectOption('320'); await settle();
  check('Groups sample compact no overflow', await settings.evaluate(e => e.scrollWidth <= e.clientWidth), true);
  await settings.screenshot({ path: path.join(output, `${label}-groups-compact.png`) });
  check('console health', consoleMessages, []); check('page errors', pageErrors, []);
})().catch(error => { failures.push(error.stack || error.message); }).finally(async () => {
  const version = browser ? browser.version() : 'not launched';
  if (browser) await browser.close();
  if (server) await new Promise(resolve => server.close(resolve));
  fs.mkdirSync(output, { recursive: true });
  fs.writeFileSync(path.join(output, `${label}.json`), JSON.stringify({ url, browser: version, boundary: process.env.STYLES_PANEL_STAGE_ONLY === '1' ? 'Actual compiled Stage, Form, sample and supplemental real Fluent probes; Browser plugin not available' : process.env.STYLES_PANEL_SHELL_ONLY === '1' ? 'Actual compiled Shell/provider/store/hooks and sample; SDK/SPFx services only; Browser plugin not available' : 'SPFx Teams/theme hook context only; actual compiled library, sample panel, React17, Griffel, FluentProvider', observations, thumbInteractionCaptures, failures, consoleMessages, pageErrors }, null, 2));
  console.log(JSON.stringify({ browser: version, checks: observations.length, failures, consoleMessages, pageErrors }, null, 2));
  if (failures.length) process.exitCode = 1;
});
