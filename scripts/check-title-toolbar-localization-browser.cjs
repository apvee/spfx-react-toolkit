// Scoped acceptance against actual source/compiled Shell and its existing host fixture.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const harness = require('../tests/fixtures/formatted-text-harness.cjs');
const root = path.resolve(__dirname, '..');
const output = path.join(root, '.docs/maintenance/title-toolbar-localization/browser');
const checks = [];
const diagnostics = [];
function check(name, actual, expected) { assert.deepEqual(actual, expected, name); checks.push(name); }
const settle = page => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
async function settlePalette(page) {
  await page.locator('[data-title-color-picker]').waitFor();
  await page.waitForFunction(() => {
    for (let node = document.querySelector('[data-title-color-picker]'); node; node = node.parentElement)
      if (node.getAnimations().some(animation => animation.playState === 'running')) return false;
    return true;
  });
}
async function serve(directory) {
  const server = http.createServer((request, response) => {
    const pathname = new URL(request.url, 'http://localhost').pathname;
    if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
    const file = path.resolve(directory, `.${pathname === '/' ? '/index.html' : pathname}`);
    if (!file.startsWith(directory + path.sep)) { response.writeHead(403).end(); return; }
    fs.readFile(file, (error, data) => {
      if (error) { response.writeHead(404).end(); return; }
      response.setHeader('Content-Type', file.endsWith('.js') ? 'text/javascript' : 'text/html'); response.end(data);
    });
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  return { url: `http://127.0.0.1:${server.address().port}`, close: () => new Promise(resolve => server.close(resolve)) };
}
if (require.main === module) (async () => {
  fs.mkdirSync(output, { recursive: true });
  const resources = Object.fromEntries(require('../tests/fixtures/sharepoint-ui-locales.json').cultures.map(culture => [culture,
    require(`../packages/spfx-react-toolkit/src/controls/web-part-shell/localization/resources/${culture}.json`)]));
  const editorResources = Object.fromEntries(Object.keys(resources).map(culture => [culture,
    require(`../packages/spfx-react-toolkit/src/controls/editable-text/localization/resources/${culture}.json`)]));
  const common = await harness.loadModule('packages/spfx-react-toolkit/src/controls/editable-text/localization/editable-text-strings.internal.ts');
  const stringsFor = culture => { const text = common.defineEditableTextStrings(editorResources[culture]); return { ...resources[culture], ...text, titleFormattingControls: text.formattingControls, moreTitleFormatting: text.moreFormatting }; };
  const dictionaryPatch = culture => ({ strings: resources[culture], titleEditorResource: editorResources[culture] });
  const { editableTextColorDescriptors: titleColorDescriptors } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/editable-text/editable-text-format.internal.ts');
  const cultures = require('../tests/fixtures/sharepoint-ui-locales.json').cultures;
  check('50 canonical dictionary catalogs', Object.keys(resources).sort(), [...cultures].sort());
  const browser = await chromium.launch({ executablePath: process.env.SX_BROWSER_EXECUTABLE || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', headless: true });
  try {
    for (const boundary of ['source', 'compiled']) {
      const server = await serve(path.join(root, 'temp', boundary === 'source' ? 'sx-browser' : 'styles-panel-browser'));
      const context = await browser.newContext({ viewport: { width: 1280, height: 1000 } });
      const page = await context.newPage();
      page.on('pageerror', error => diagnostics.push(`${boundary}: ${error.message}`));
      page.on('console', message => { if (message.type() === 'error') diagnostics.push(`${boundary}: ${message.text()}`); });
      try {
        await page.goto(server.url);
        const render = async patch => { await page.evaluate(value => window.shellFixture.render(value), patch); await settle(page); };
        await render({ mode: 'Edit', ...dictionaryPatch('en-US'), otherStrings: resources['it-IT'], otherTitleEditorResource: editorResources['it-IT'], width: 900,
          properties: { title: 'Author content', titleHTML: '<h2><strong>Author content</strong></h2>', showTitle: true } });
        const editable = page.locator('[data-shell-root=first] [contenteditable=true]'); await editable.waitFor();
        const identity = await editable.elementHandle();
        const original = await page.evaluate(() => window.shellFixture.observe());
        const toolbar = page.locator('[data-title-toolbar]');
        for (const culture of process.env.TITLE_LOCALIZATION_SKIP_CATALOGS === '1' ? [] : cultures) {
          const strings = stringsFor(culture);
          await render(dictionaryPatch(culture));
          check(`${boundary}/${culture} editor identity`, await editable.evaluate((element, original) => element === original, identity), true);
          check(`${boundary}/${culture} editor accessible name`, await editable.getAttribute('aria-label'), strings.titleEditorLabel);
          await editable.focus(); await toolbar.waitFor();
          check(`${boundary}/${culture} toolbar name`, await toolbar.getAttribute('aria-label'), strings.titleFormattingControls);
          check(`${boundary}/${culture} heading selector name`, await toolbar.getByRole('combobox', { name: strings.heading, exact: true }).count(), 1);
          check(`${boundary}/${culture} size selector name`, await toolbar.getByRole('combobox', { name: strings.fontSize, exact: true }).count(), 1);
          check(`${boundary}/${culture} size expand button name`, await toolbar.locator('.fui-Combobox__expandIcon').getAttribute('aria-label'), strings.fontSize);
          for (const key of ['bold', 'italic', 'underline', 'clearFormatting', 'spacing'])
            check(`${boundary}/${culture} ${key} name`, await toolbar.getByRole('button', { name: strings[key], exact: true }).count(), 1);
          await toolbar.getByRole('button', { name: strings.themeColor, exact: true }).click();
          const palette = page.locator('[data-title-color-picker]'); await settlePalette(page);
          const layout = await palette.evaluate(element => {
            const sections = ['automatic', 'theme', 'palette'].map(id => element.querySelector(`[data-title-color-section=${id}]`));
            if (sections.some(section => !section)) return { stacked: false, horizontal: false };
            const rects = sections.map(section => section.getBoundingClientRect());
            return {
              stacked: rects.every((rect, index) => index === 0 || rect.top >= rects[index - 1].bottom - 1),
              horizontal: sections.slice(1).every(section => {
                const swatches = [...section.querySelectorAll('[data-title-color]')];
                if (swatches.length < 2) return true;
                const [first, second] = swatches.map(swatch => swatch.getBoundingClientRect());
                return Math.abs(first.top - second.top) <= 1 && Math.abs(first.left - second.left) > 1;
              }),
            };
          });
          check(`${boundary}/${culture} palette sections stack vertically`, layout.stacked, true);
          check(`${boundary}/${culture} color rows flow horizontally`, layout.horizontal, true);
          for (const key of ['automatic', 'themeColors', 'paletteColors'])
            check(`${boundary}/${culture} ${key} group`, await palette.getByText(strings[key], { exact: true }).count(), 1);
          const swatches = palette.locator('[data-title-color]');
          const ids = await swatches.evaluateAll(elements => elements.map(element => element.getAttribute('data-title-color')));
          check(`${boundary}/${culture} automatic is present`, ids.includes('automatic'), true);
          check(`${boundary}/${culture} no custom colors`, ids.every(id => id === 'automatic' || titleColorDescriptors.some(item => item.id === id)), true);
          for (const id of ids) {
            const name = id === 'automatic' ? strings.automatic : strings.colorNames[id];
            const swatch = palette.locator(`[data-title-color="${id}"]`);
            check(`${boundary}/${culture}/${id} accessible color name`, await swatch.getAttribute('aria-label'), name);
            await swatch.hover();
            const tooltip = page.getByRole('tooltip').filter({ hasText: name }); await tooltip.waitFor({ state: 'visible' });
            check(`${boundary}/${culture}/${id} hover tooltip`, await tooltip.textContent(), name);
          }
          await page.keyboard.press('Escape'); await palette.waitFor({ state: 'hidden' });
          const controlBoxes = await toolbar.evaluate(element => {
            const boxes = [...element.querySelectorAll('.fui-Dropdown,.fui-Combobox,.fui-Button')]
              .filter(control => !control.closest('[data-overflowing]') && control.getBoundingClientRect().width > 0)
              .map(control => { const r = control.getBoundingClientRect(); return { height: r.height, center: r.top + r.height / 2, selector: control.matches('.fui-Dropdown,.fui-Combobox') }; });
            return { selectorsSmall: boxes.filter(box => box.selector).every(box => Math.round(box.height) === 24), centered: boxes.length > 0 && boxes.every(box => Math.abs(box.center - boxes[0].center) <= 1) };
          });
          check(`${boundary}/${culture} selectors are small`, controlBoxes.selectorsSmall, true);
          check(`${boundary}/${culture} controls vertically centered`, controlBoxes.centered, true);
          if (['en-US', 'it-IT', 'de-DE', 'ar-SA', 'he-IL', 'prs-AF', 'zh-CN', 'sr-Cyrl-RS'].includes(culture)) {
            await editable.focus(); await toolbar.getByRole('button', { name: strings.themeColor, exact: true }).click();
            const first = palette.locator('[data-title-color=automatic]');
            await page.mouse.move(0, 0);
            // The popup initially focuses Automatic. Move away and back using
            // native SwatchPicker keys so this tests a real new focus event.
            await first.press('ArrowRight'); await page.keyboard.press('ArrowLeft');
            check(`${boundary}/${culture} keyboard returns to Automatic`, await first.evaluate(element => document.activeElement === element), true);
            const focusTooltip = page.getByRole('tooltip').filter({ hasText: strings.automatic }); await focusTooltip.waitFor({ state: 'visible' });
            check(`${boundary}/${culture} focus tooltip`, await focusTooltip.textContent(), strings.automatic);
            await settlePalette(page);
            await page.screenshot({ path: path.join(output, `${boundary}-${culture}-palette.png`) });
            await page.keyboard.press('Escape'); await palette.waitFor({ state: 'hidden' });
          }
          if ((cultures.indexOf(culture) + 1) % 10 === 0) console.log(`${boundary}: ${cultures.indexOf(culture) + 1}/50 dictionary cultures checked.`);
        }
        check(`${boundary} dictionary changes do not publish`, (await page.evaluate(() => window.shellFixture.observe())).refreshes, original.refreshes);
        check(`${boundary} author content remains literal`, await editable.textContent(), 'Author content');
        check(`${boundary} second instance dictionary independent`, await page.locator('[data-shell-root=second] [contenteditable=true]').getAttribute('aria-label'), resources['it-IT'].titleEditorLabel);
        await render(dictionaryPatch('en-US'));
        const english = stringsFor('en-US');
        await editable.focus(); await page.keyboard.press('Meta+a');
        await toolbar.getByRole('button', { name: english.themeColor, exact: true }).click();
        await page.locator('[data-title-color=red]').click();
        await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML.includes('wps-color-red'));
        const colored = await page.evaluate(() => window.shellFixture.observe());
        await toolbar.getByRole('button', { name: english.themeColor, exact: true }).click(); await page.locator('[data-title-color=automatic]').click();
        check(`${boundary} Automatic removes only color`, await page.evaluate(() => window.shellFixture.observe().properties.titleHTML), '<h2><strong>Author content</strong></h2>');
        await editable.focus(); await page.keyboard.press('Meta+z');
        await page.waitForFunction(value => window.shellFixture.observe().properties.titleHTML === value, colored.properties.titleHTML);
        for (const rtl of [false, true]) for (const width of [1280, 640, 320]) {
          await page.setViewportSize({ width, height: 1000 }); await render({ width: width <= 640 ? width - 24 : 900, dir: rtl ? 'rtl' : 'ltr' });
          await editable.focus(); await toolbar.waitFor();
          // Fluent commits overflow visibility after its resize microtask and
          // layout effects; don't act on the mask from the previous viewport.
          await page.waitForTimeout(350);
          console.log(`${boundary}: geometry ${rtl ? 'RTL' : 'LTR'} ${width}px.`);
          const more = toolbar.getByRole('button', { name: english.moreTitleFormatting, exact: true });
          const direct = toolbar.getByRole('button', { name: english.themeColor, exact: true });
          if (await direct.isVisible()) await direct.click(); else { await more.click(); await page.getByRole('menuitem', { name: english.themeColor, exact: true }).click(); }
          const palette = page.locator('[data-title-color-picker]'); await settlePalette(page);
          const fits = await palette.evaluate(element => { const r = element.getBoundingClientRect(); return r.left >= -1 && r.right <= innerWidth + 1 && element.scrollWidth <= element.clientWidth + 1; });
          check(`${boundary}/${rtl}/${width} palette fits`, fits, true);
          await page.screenshot({ path: path.join(output, `${boundary}-${rtl ? 'rtl' : 'ltr'}-${width}.png`) });
          await page.keyboard.press('Escape'); await palette.waitFor({ state: 'hidden' });
          check(`${boundary}/${rtl}/${width} icon buttons same height`, await toolbar.locator('.fui-Button').evaluateAll(elements => {
            const heights = elements.filter(element => element.getClientRects().length && !element.closest('[data-overflowing]')).map(element => Math.round(element.getBoundingClientRect().height));
            return heights.length > 0 && new Set(heights).size === 1;
          }), true);
        }
        await identity.dispose();
        console.log(`${boundary}: dictionary/tooltips/selection/geometry checks passed.`);
      } catch (error) {
        await page.screenshot({ path: path.join(output, `${boundary}-failure.png`), fullPage: true });
        fs.writeFileSync(path.join(output, `${boundary}-failure.html`), await page.content());
        throw error;
      } finally { await context.close(); await server.close(); }
    }
    check('Browser diagnostics', diagnostics, []);
    console.log(`${checks.length} title toolbar localization browser checks PASS.`);
  } finally {
    fs.writeFileSync(path.join(output, 'report.json'), JSON.stringify({ checks, diagnostics, authenticatedSharePoint: 'NOT EXECUTED' }, undefined, 2));
    await browser.close(); await harness.cleanup();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
module.exports = { serve, settle, settlePalette };
