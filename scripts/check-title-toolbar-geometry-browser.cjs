// Native geometry, keyboard and existing Shell acceptance; no dependency installs.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const { serve, settle, settlePalette } = require('./check-title-toolbar-localization-browser.cjs');
const { checkWebPartShellBrowser } = require('./web-part-shell-browser.helpers.cjs');
const root = path.resolve(__dirname, '..');
const output = path.join(root, '.docs/maintenance/title-toolbar-localization/geometry');
const checks = [];
const diagnostics = [];
const executablePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const check = (name, actual, expected) => { assert.deepEqual(actual, expected, name); checks.push(name); };
const paletteForegroundTokens = {
  darkRed: 'colorPaletteRedForeground1', amber: 'colorPaletteMarigoldForeground3',
  yellowGreen: { first: 'colorPaletteYellowForeground3', second: 'colorPaletteGreenForeground3', percent: 73 },
  cyan: 'colorPaletteLightTealForeground2', navy: 'colorPaletteNavyForeground2',
  red: 'colorPaletteRedForeground3', orange: 'colorPaletteDarkOrangeForeground3',
  yellow: 'colorPaletteYellowForeground3', green: 'colorPaletteGreenForeground3',
  teal: 'colorPaletteTealForeground2', blue: 'colorPaletteBlueForeground2',
  purple: 'colorPalettePurpleForeground2', pink: 'colorPalettePinkForeground2', brown: 'colorPaletteBrownForeground2',
};
function hexToRgb(value) {
  const match = /^#([\da-f]{6})$/i.exec(value);
  assert.ok(match, 'standard Fluent light/dark foreground is a six-digit hex value');
  const digits = match[1];
  return `rgb(${[0, 2, 4].map(index => parseInt(digits.slice(index, index + 2), 16)).join(', ')})`;
}
async function palette(page) {
  const toolbar = page.locator('[data-title-toolbar]');
  const direct = toolbar.getByRole('button', { name: 'Theme color', exact: true });
  if (await direct.isVisible()) await direct.click();
  else { await toolbar.getByRole('button', { name: 'More text formatting', exact: true }).click(); await page.getByRole('menuitem', { name: 'Theme color', exact: true }).click(); }
  await settlePalette(page);
}
async function checkColorOverflowResize(page, editable, boundary) {
  const toolbar = page.locator('[data-title-toolbar]');
  for (const dir of ['ltr', 'rtl']) {
    await page.evaluate(dir => window.shellFixture.render({ width: 296, dir, theme: 'light' }), dir);
    await editable.focus(); await settle(page); await page.waitForTimeout(350);
    await palette(page);
    // Resize only: no focus changes, property updates or synthetic dismissal.
    await page.evaluate(() => window.shellFixture.render({ width: 500 }));
    await settle(page); await page.waitForTimeout(350);
    check(`${boundary}/${dir} color becomes direct while More remains`,
      await toolbar.getByRole('button', { name: 'Theme color', exact: true }).isVisible()
        && await toolbar.getByRole('button', { name: 'More text formatting', exact: true }).isVisible(), true);
    check(`${boundary}/${dir} removed overflow palette closes`, await page.locator('[data-title-color-picker]').count(), 0);
    await page.evaluate(() => window.shellFixture.render({ width: 296 }));
    await settle(page); await page.waitForTimeout(350);
    check(`${boundary}/${dir} narrowing does not reopen palette`, await page.locator('[data-title-color-picker]').count(), 0);
    check(`${boundary}/${dir} narrowing starts with closed More`, await page.getByRole('menuitem', { name: 'Theme color', exact: true }).count(), 0);
    await palette(page);
    check(`${boundary}/${dir} palette can be reopened deliberately`, await page.locator('[data-title-color-picker]').count(), 1);
    await page.keyboard.press('Escape'); await page.keyboard.press('Escape');
    await page.evaluate(() => window.shellFixture.render({ width: 500 }));
    await editable.focus(); await settle(page); await page.waitForTimeout(350);
    await palette(page);
    await page.evaluate(() => window.shellFixture.render({ width: 296 }));
    await settle(page); await page.waitForTimeout(350);
    check(`${boundary}/${dir} direct palette closes when its trigger hides`, await page.locator('[data-title-color-picker]').count(), 0);
    await page.evaluate(() => window.shellFixture.render({ width: 500 }));
    await settle(page); await page.waitForTimeout(350);
    check(`${boundary}/${dir} widening does not reopen direct palette`, await page.locator('[data-title-color-picker]').count(), 0);
    await editable.focus(); await palette(page);
    check(`${boundary}/${dir} direct palette can be reopened deliberately`, await page.locator('[data-title-color-picker]').count(), 1);
    await page.keyboard.press('Escape');
  }
}
(async () => {
  fs.mkdirSync(output, { recursive: true });
  const browser = await chromium.launch({ executablePath, headless: true });
  try {
    for (const boundary of ['source', 'compiled']) {
      const server = await serve(path.join(root, 'temp', boundary === 'source' ? 'sx-browser' : 'styles-panel-browser'));
      try {
        const page = await browser.newPage({ viewport: { width: 1280, height: 1000 } });
        page.on('pageerror', error => diagnostics.push(error.message));
        await page.goto(server.url);
        if (process.env.TITLE_TOOLBAR_SKIP_LEGACY !== '1' && process.env.TITLE_TOOLBAR_RESET_ONLY !== '1') {
          await checkWebPartShellBrowser(page, { url: server.url, output, label: boundary, check, consoleMessages: [] });
          console.log(`${boundary}: existing Shell/native title acceptance passed.`);
        }
        await page.evaluate(() => window.shellFixture.render({ mode: 'Edit', uiCulture: 'en-US', width: 500, properties: { title: 'Themed title', titleHTML: '<h2>Themed title</h2>', showTitle: true } }));
        const editable = page.locator('[data-shell-root=first] [contenteditable=true]'); await editable.waitFor();
        await checkColorOverflowResize(page, editable, boundary);
        if (process.env.TITLE_TOOLBAR_RESET_ONLY === '1') { await page.close(); continue; }
        await page.evaluate(() => window.shellFixture.render({ width: 500, dir: 'ltr' }));
        for (const theme of ['light', 'dark', 'highContrast']) {
          await page.evaluate(theme => window.shellFixture.render({ theme }), theme); await editable.focus(); await page.keyboard.press('Meta+a'); await settle(page); await palette(page);
          const offeredColors = await page.locator('[data-title-color-section=palette]').evaluate((element, mapping) =>
            [...element.querySelectorAll('[data-title-color]')].map(swatch => {
              const id = swatch.getAttribute('data-title-color'); const css = getComputedStyle(swatch);
              const source = mapping[id];
              const token = name => css.getPropertyValue(`--${name}`).trim();
              const expected = typeof source === 'string' ? token(source)
                : `color-mix(in srgb, ${token(source.first)} ${source.percent}%, ${token(source.second)})`;
              return { id, actual: css.getPropertyValue('--fui-SwatchPicker--color').trim(), expected };
            }), paletteForegroundTokens);
          if (theme === 'light') {
            check(`${boundary} reference Theme group has12 ordered swatches`, await page.locator('[data-title-color-section=theme] [data-title-color]').evaluateAll(elements => elements.map(element => element.getAttribute('data-title-color'))), ['primary','neutralDark','subtle','neutralMuted','neutralLight','neutralLighter','brandDarkest','brandDarker','brandDark','brandAlt','brand','brandLight']);
            check(`${boundary} reference Standard group has10 ordered swatches`, offeredColors.map(color => color.id), ['darkRed','red','amber','yellow','yellowGreen','green','cyan','blue','navy','purple']);
          }
          check(`${boundary}/${theme} offers named palette foregrounds`, offeredColors.length > 0, true);
          for (const color of offeredColors) check(`${boundary}/${theme}/${color.id} uses correct foreground`, color.actual, color.expected);
          const tokenValues = await page.locator('[data-title-color-picker]').evaluate(element => {
            const css = getComputedStyle(element);
            return ['colorNeutralForeground1', 'colorBrandForeground1', 'colorPaletteRedForeground3'].map(name => css.getPropertyValue(`--${name}`).trim());
          });
          check(`${boundary}/${theme} inherits all foreground token families`, tokenValues.every(Boolean), true);
          await page.screenshot({ path: path.join(output, `${boundary}-${theme}.png`) });
          if (theme === 'light' || theme === 'dark') {
            const yellow = offeredColors.find(color => color.id === 'yellow'); assert.ok(yellow, 'yellow is offered in standard Fluent themes');
            check(`${boundary}/${theme} yellow swatch displays its foreground hue`, await page.locator('[data-title-color=yellow]').evaluate(element => getComputedStyle(element).backgroundColor), hexToRgb(yellow.expected));
            await page.locator('[data-title-color=yellow]').click();
            await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML.includes('wps-color-yellow'));
            check(`${boundary}/${theme} yellow title uses the same foreground hue`, await editable.locator('.wps-color-yellow').evaluate(element => getComputedStyle(element).color), hexToRgb(yellow.expected));
            await editable.focus(); await palette(page);
            const yellowGreen = page.locator('[data-title-color=yellowGreen]');
            const mixedHue = await yellowGreen.evaluate(element => getComputedStyle(element).backgroundColor);
            await yellowGreen.click();
            await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML.includes('wps-color-yellowGreen'));
            check(`${boundary}/${theme} mixed foreground title matches its swatch`, await editable.locator('.wps-color-yellowGreen').evaluate(element => getComputedStyle(element).color), mixedHue);
          } else await page.keyboard.press('Escape');
        }
        await page.emulateMedia({ forcedColors: 'active' }); await editable.focus(); await palette(page);
        check(`${boundary} forced colors keeps Automatic named`, await page.locator('[data-title-color=automatic]').getAttribute('aria-label'), 'Automatic');
        check(`${boundary} forced colors keeps offered colors named`, await page.locator('[data-title-color]').evaluateAll(elements => elements.length > 1 && elements.every(element => Boolean(element.getAttribute('aria-label')?.trim()))), true);
        await page.screenshot({ path: path.join(output, `${boundary}-forced-colors.png`) }); await page.keyboard.press('Escape'); await page.emulateMedia({ forcedColors: 'none' });
        await page.close();
        const metrics = [];
        for (const zoom of [1, 2]) {
          const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'title-toolbar-native-zoom-'));
          fs.mkdirSync(path.join(profile, 'Default'));
          fs.writeFileSync(path.join(profile, 'Default/Preferences'), JSON.stringify({ partition: { default_zoom_level: { x: Math.log(zoom) / Math.log(1.2) } } }));
          let context;
          try {
            context = await chromium.launchPersistentContext(profile, { executablePath, headless: true, viewport: null, args: ['--window-size=1280,1000'] });
            const native = context.pages()[0]; await native.goto(server.url);
            await native.evaluate(() => window.shellFixture.render({ mode: 'Edit', width: 500, properties: { title: 'Zoom title', titleHTML: '<h2>Zoom title</h2>', showTitle: true } }));
            const editor = native.locator('[data-shell-root=first] [contenteditable=true]'); await editor.waitFor();
            for (const dir of ['ltr', 'rtl']) {
              await native.evaluate(dir => window.shellFixture.render({ dir }), dir); await editor.focus(); await settle(native); await palette(native);
              check(`${boundary}/${zoom}/${dir} palette stays within viewport`, await native.locator('[data-title-color-picker]').evaluate(element => { const r = element.getBoundingClientRect(); return r.left >= -1 && r.right <= innerWidth + 1 && element.scrollWidth <= element.clientWidth + 1; }), true);
              const session = await context.newCDPSession(native);
              const shot = await session.send('Page.captureScreenshot', { format: 'png', fromSurface: true, captureBeyondViewport: false });
              fs.writeFileSync(path.join(output, `${boundary}-${zoom * 100}-${dir}.png`), Buffer.from(shot.data, 'base64')); await session.detach();
              await native.keyboard.press('Escape');
            }
            await native.goto(`${server.url}/?scenario=editable-text`);
            await native.locator('[data-editable-text-demo=mode]').selectOption('edit');
            const standalone = native.locator('[data-editable-text-demo=preview] [contenteditable=true]');
            await standalone.waitFor(); await standalone.focus(); await settle(native);
            check(`${boundary}/${zoom} standalone toolbar fits native zoom`, await native.locator('[data-title-toolbar]').evaluate(element => { const rect = element.getBoundingClientRect(); return rect.left >= -1 && rect.right <= innerWidth + 1; }), true);
            metrics.push(await native.evaluate(() => ({ dpr: devicePixelRatio, width: innerWidth })));
          } finally { await context?.close(); fs.rmSync(profile, { recursive: true, force: true }); }
        }
        check(`${boundary} 200% is native double DPR`, metrics[1].dpr / metrics[0].dpr, 2);
        check(`${boundary} 200% halves CSS viewport`, metrics[0].width / metrics[1].width, 2);
      } finally { await server.close(); }
    }
    check('Scoped browser runtime errors', diagnostics, []);
    console.log(`${checks.length} native title geometry/legacy checks PASS.`);
  } finally { fs.writeFileSync(path.join(output, 'report.json'), JSON.stringify({ checks, diagnostics, authenticatedSharePoint: 'NOT EXECUTED' }, undefined, 2)); await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
