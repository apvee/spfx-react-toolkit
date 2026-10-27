// Actual CardsPanel, independently bundled against source and compiled library.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const root = path.resolve(__dirname, '..');
const evidence = path.resolve(process.env.CARDS_BROWSER_EVIDENCE || path.join(root, '.docs/maintenance/evidence/cards'));
async function verify(boundary) {
  const directory = path.join(root, 'temp', boundary === 'source' ? 'sx-browser' : 'styles-panel-browser');
  const server = http.createServer((request, response) => {
    const pathname = new URL(request.url, 'http://127.0.0.1').pathname;
    if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
    // Real decode failure, without an unrelated HTTP console error.
    if (pathname === '/cards-demo-missing-image.png') { response.writeHead(200, { 'Content-Type': 'image/png' }).end('invalid image'); return; }
    const file = { '/': 'index.html', '/fixture.js': 'fixture.js', '/fixture.js.map': 'fixture.js.map' }[pathname];
    if (!file) { response.writeHead(404).end(); return; }
    response.setHeader('Content-Type', file.endsWith('.html') ? 'text/html' : 'application/javascript');
    response.end(fs.readFileSync(path.join(directory, file)));
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const checks = [], errors = [], warnings = [];
  const check = (name, actual, expected) => { assert.deepEqual(actual, expected, name); checks.push(name); };
  let browser;
  try {
    browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
    const page = await browser.newPage({ viewport: { width: 1280, height: 1100 } });
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); if (message.type() === 'warning') warnings.push(message.text()); });
    const url = `http://127.0.0.1:${server.address().port}`;
    await page.goto(url);
    const live = page.locator('[data-card-instance="main"]');
    const liveTarget = marker => live.locator(`[data-cards-demo="${marker}"]`);
    const liveSubject = () => live.locator('[data-cards-demo="preview-card"], [data-cards-demo="content-card"]');
    check('product has one configuration', await live.locator('[data-demo-configuration]').count(), 1);
    check('product has one actual card', await live.locator('.fui-Card').count(), 2);
    check('approved title', await liveTarget('title-input').inputValue(), 'Quarterly report');
    check('approved description', await liveTarget('description-input').inputValue(), 'Team results, goals and next steps.');
    check('approved orientation', await liveTarget('orientation').inputValue(), 'horizontal');
    await live.screenshot({ path: path.join(evidence, `${boundary}-single-wide.png`) });
    await page.setViewportSize({ width: 390, height: 1100 });
    await live.screenshot({ path: path.join(evidence, `${boundary}-single-narrow.png`) });
    check('narrow product fits', await liveTarget('preview-container').evaluate(node => node.getBoundingClientRect().right <= window.innerWidth), true);
    await page.setViewportSize({ width: 1280, height: 1100 });
    await liveTarget('retained-input').fill('Single example retained note');
    await liveTarget('retained-input').evaluate(node => { window.liveRetainedInput = node; });
    await liveTarget('orientation').selectOption('vertical'); await liveTarget('width').selectOption('compact');
    await liveTarget('loading').check(); await liveTarget('loading').uncheck();
    check('single example keeps note DOM through layout and loading', await liveTarget('retained-input').evaluate(node => node === window.liveRetainedInput && node.value === 'Single example retained note'), true);
    await live.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click();
    check('open diagnostics still has one actual card', await live.locator('.fui-Card').count(), 2);
    await liveTarget('kind').selectOption('content');
    check('alternate card is a single ContentCard', await liveTarget('content-card').count(), 1);
    check('alternate removes PreviewCard', await liveTarget('preview-card').count(), 0);
    check('alternate preserves shared note', await liveTarget('retained-input').inputValue(), 'Single example retained note');
    await liveTarget('kind').selectOption('preview');
    check('switch back has one subject', await liveSubject().count(), 1);
    const liveMediaBox = () => liveTarget('preview-card').evaluate(node => { const r = node.querySelector('[data-card-preview]').getBoundingClientRect(); return { width: r.width, height: r.height }; });
    const originalLiveMedia = await liveMediaBox();
    await liveTarget('fail-image').click(); await liveTarget('preview-card').locator('[data-card-image-fallback]').waitFor();
    check('single example fallback keeps media geometry', await liveMediaBox(), originalLiveMedia);
    await liveTarget('new-src').click();
    check('single example new source clears fallback', await liveTarget('preview-card').locator('[data-card-image-fallback]').count(), 0);
    await liveTarget('width').selectOption('standard'); await liveTarget('orientation').selectOption('horizontal');
    await live.getByRole('button', { name: 'Advanced diagnostics', exact: true }).click();

    await page.getByRole('button', { name: 'Independent RTL instance', exact: true }).click();
    await page.locator('[data-card-instance=regression]').getByRole('button', { name: 'Selection and native Card comparison', exact: true }).click();
    const main = page.locator('[data-card-instance="regression"]'), independent = page.locator('[data-card-instance="independent"]');
    const diagnostics = main.getByRole('button', { name: 'Advanced diagnostics', exact: true });
    if (await diagnostics.getAttribute('aria-expanded') !== 'true') await diagnostics.click();
    const target = name => main.locator(`[data-cards-demo="${name}"]`);
    const panel = page.locator('[data-cards-demo="regression-panel"]');
    const preview = target('preview-card'), content = target('content-card'), native = target('native-card');
    // A second Fluent bundle can insert its native full-bleed preview rules later,
    // as in the SharePoint workbench alongside another web part. Reuse the installed
    // peer's actual selectors rather than depending on generated class names.
    const lateNativePreviewRules = await page.evaluate(() => {
      const rules = [];
      function collect(list) {
        for (const rule of list) {
          if (rule.selectorText?.includes('.fui-CardPreview') && /margin/.test(rule.style?.cssText || '')) rules.push(rule.cssText);
          if (rule.cssRules) collect(rule.cssRules);
        }
      }
      for (const sheet of document.styleSheets) collect(sheet.cssRules);
      return rules.join('\n');
    });
    assert.ok(lateNativePreviewRules, 'installed native preview margin rules are present');
    await page.addStyleTag({ content: lateNativePreviewRules });
    const settle = () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    const counters = () => target('counters').textContent();
    const inspect = async () => { await target('inspect-refs').click(); return target('observations').textContent(); };
    const box = () => preview.evaluate(node => {
      const media = node.querySelector('[data-card-preview]').getBoundingClientRect();
      const detail = node.querySelector('[data-card-detail]').getBoundingClientRect();
      const css = getComputedStyle(node);
      return { wrapped: detail.top >= media.bottom - 1, alongside: Math.abs(detail.top - media.top) < 1, fits: node.scrollWidth <= node.clientWidth + 1,
        mediaWidth: media.width, mediaHeight: media.height, padding: css.paddingTop, direction: css.direction };
    });
    const assertPreviewInsets = async name => {
      const insets = await preview.evaluate(node => {
        const rootBox = node.getBoundingClientRect();
        const media = node.querySelector('[data-card-preview]');
        const mediaBox = media.getBoundingClientRect();
        const rootCss = getComputedStyle(node);
        const mediaCss = getComputedStyle(media);
        const share = node.querySelector('[data-cards-demo="header-action"]');
        const shareLabel = Array.from(share.childNodes).find(child => child.nodeType === Node.TEXT_NODE && child.textContent.trim() === 'Share');
        const labelRange = document.createRange();
        labelRange.selectNodeContents(shareLabel);
        const inlineInset = rootCss.direction === 'rtl' ? rootBox.right - mediaBox.right : mediaBox.left - rootBox.left;
        const inlinePadding = parseFloat(rootCss.direction === 'rtl' ? rootCss.paddingRight : rootCss.paddingLeft);
        return {
          top: Math.abs(mediaBox.top - rootBox.top - parseFloat(rootCss.paddingTop)) < 1,
          inline: Math.abs(inlineInset - inlinePadding) < 1,
          margins: [mediaCss.marginTop, mediaCss.marginRight, mediaCss.marginBottom, mediaCss.marginLeft],
          shareLabelLines: labelRange.getClientRects().length,
        };
      });
      check(`${name} preserves top padding after late native CSS`, insets.top, true);
      check(`${name} preserves inline padding after late native CSS`, insets.inline, true);
      check(`${name} has no full-bleed margins`, insets.margins, ['0px', '0px', '0px', '0px']);
      check(`${name} Share label remains on one line`, insets.shareLabelLines, 1);
    };
    check('title', await page.title(), `Cards ${boundary} browser checkpoint`);
    check('boundary', await page.locator('[data-cards-fixture="boundary"]').textContent(), boundary);
    check('page identity', new URL(page.url()).origin, url);
    check('native root div', await preview.evaluate(node => node.tagName), 'DIV');
    check('owned heading level', await preview.locator('h2').textContent(), '**Literal report title**');
    check('default content heading', await content.locator('h3').textContent(), '**Literal content title**');
    check('literal owned title', await preview.locator('h2 strong').count(), 0);
    check('explicit markdown composition', await preview.locator('[data-card-body] strong').textContent(), 'FormattedText');
    check('slot ref hierarchy', (await inspect()).includes('root:DIV title:SPAN preview:absent logo:IMG button:BUTTON'), true);
    check('consumer title class', await target('title-slot').evaluate(node => node.classList.contains('cards-consumer-title')), true);
    await page.waitForFunction(() => [...document.querySelectorAll('[data-card-instance="regression"] [data-card-preview] img')].every(image => image.complete && image.naturalWidth > 0));
    check('vertical media ratio', Math.abs((await box()).mediaWidth / (await box()).mediaHeight - 16 / 9) < 0.02, true);
    await assertPreviewInsets('vertical');
    check('logo inner bounds', await target('logo-slot').evaluate(node => { const r = node.getBoundingClientRect(); return r.width <= 32 && r.height <= 32 && r.width > 0; }), true);
    check('image logo is a loaded square icon resource', await target('logo-slot').evaluate(node => node.complete && node.naturalWidth > 0 && node.naturalWidth === node.naturalHeight), true);
    const input = target('retained-input');
    await input.fill('Retained caller edit');
    await input.evaluate(node => { window.cardsRetainedInput = node; });
    await target('orientation').check(); await settle();
    check('wide horizontal alongside', (await box()).alongside, true);
    await assertPreviewInsets('wide horizontal');
    check('native Card medium padding', (await box()).padding, await native.evaluate(node => getComputedStyle(node).paddingTop));
    check('horizontal media preferred width', (await box()).mediaWidth >= 128, true);
    await preview.screenshot({ path: path.join(evidence, `${boundary}-horizontal-wide.png`) });
    await target('width').check(); await settle();
    check('narrow horizontal wraps by fit', (await box()).wrapped, true);
    check('narrow preview fits', (await box()).fits, true);
    check('wrapped media ratio', Math.abs((await box()).mediaWidth / (await box()).mediaHeight - 16 / 9) < 0.02, true);
    await assertPreviewInsets('wrapped horizontal');
    await preview.screenshot({ path: path.join(evidence, `${boundary}-horizontal-narrow.png`) });
    await page.evaluate(() => window.demoFixtureHost.update({ theme: 'dark' }));
    await panel.locator('[data-cards-demo="direction"]').check(); await settle();
    check('RTL direction', (await box()).direction, 'rtl');
    await assertPreviewInsets('dark RTL');
    check('state and native node retained', await input.evaluate(node => window.cardsRetainedInput === node && node.value === 'Retained caller edit'), true);
    check('independent instance width', await independent.locator('[data-cards-demo=preview-container]').evaluate(node => Math.round(node.getBoundingClientRect().width)), 520);
    check('independent instance orientation', await independent.locator('[data-cards-demo="orientation"]').isChecked(), false);
    await preview.screenshot({ path: path.join(evidence, `${boundary}-dark-rtl.png`) });
    await target('header-action').click(); await target('footer-action').click();
    check('native child click handlers and bubbling', (await counters()).startsWith('Header:1 Footer:1 Root bubbled:2'), true);
    await target('disabled').check();
    check('disabled Card keeps child native disabled flag off', await target('footer-action').evaluate(node => node.disabled), false);
    // Native Card applies ancestor aria-disabled; test actual pointer dispatch instead of Playwright's inherited disabled heuristic.
    await target('footer-action').scrollIntoViewIfNeeded();
    const disabledChildBox = await target('footer-action').boundingBox();
    await page.mouse.click(disabledChildBox.x + disabledChildBox.width / 2, disabledChildBox.y + disabledChildBox.height / 2);
    check('native disabled overlay blocks pointer activation', (await counters()).startsWith('Header:1 Footer:1'), true);
    await target('disabled').uncheck();
    for (const key of ['Enter', 'Space']) {
      const before = (await content.getAttribute('aria-checked')) === 'true';
      const nativeBefore = (await native.getAttribute('aria-checked')) === 'true';
      await content.focus(); await page.keyboard.press(key);
      await native.focus(); await page.keyboard.press(key);
      check(`native root ${key} selection parity`, (await content.getAttribute('aria-checked')) !== String(before), (await native.getAttribute('aria-checked')) !== String(nativeBefore));
    }
    // Pin the installed native peer's descendant keyboard behavior.
    const beforeChild = await content.getAttribute('aria-checked'), beforeNativeChild = await native.getAttribute('aria-checked');
    await target('selection-child').focus(); await page.keyboard.press('Enter');
    await target('native-child').focus(); await page.keyboard.press('Enter');
    check('native child Enter selection parity', (await content.getAttribute('aria-checked')) !== beforeChild, (await native.getAttribute('aria-checked')) !== beforeNativeChild);
    await target('preview-mode').selectOption('custom');
    check('custom preview image overrides generated image', await preview.getByRole('img', { name: 'Caller-owned custom cover' }).count(), 1);
    check('custom preview inner ref', (await inspect()).includes('preview:DIV'), true);
    await target('logo-callback').check();
    check('bounded logo callback ref', (await inspect()).includes('logo:DIV'), true);
    check('callback logo renders a Fluent SVG', await target('logo-slot').locator('svg path[d]').count() > 0, true);
    for (const mode of ['empty-object', 'empty-callback']) {
      await target('preview-mode').selectOption(mode);
      check(`${mode} requests owned region`, await preview.locator('[data-card-preview]').count(), 1);
      check(`${mode} suppresses generated image`, await preview.getByRole('img', { name: 'White circle on a colored report cover' }).count(), 0);
      check(`${mode} retains logo`, await target('logo-slot').count(), 1);
    }
    for (const mode of ['false', 'empty-fragment']) {
      await target('preview-mode').selectOption(mode);
      check(`${mode} suppresses region and logo`, await preview.locator('[data-card-preview], [data-cards-demo="logo-slot"]').count(), 0);
      const first = await inspect();
      await target('custom-fallback').check();
      const second = await inspect();
      const count = value => value.match(/preview callbacks:(\d+) fallback callbacks:(\d+) logo callbacks:(\d+)/)[0];
      check(`${mode} does not evaluate suppressed callbacks`, count(second), count(first));
      await target('custom-fallback').uncheck();
    }
    await target('preview-mode').selectOption('image');
    const mediaGeometry = () => preview.evaluate(node => {
      const media = node.querySelector('[data-card-preview]');
      const failed = media.querySelector('[data-card-image-fallback]');
      const surface = failed ? failed.firstElementChild || failed : media.querySelector(':scope > img');
      const size = element => { const rect = element.getBoundingClientRect(); return { width: rect.width, height: rect.height }; };
      return { card: size(node), preview: size(media), surface: size(surface) };
    });
    const loadedGeometry = await mediaGeometry();
    await target('fail-image').click();
    await preview.locator('[data-card-image-fallback]').waitFor();
    check('default fallback fills the same viewport as the loaded image', await mediaGeometry(), loadedGeometry);
    check('failed media accessible name', await preview.locator('[data-card-image-fallback]').getAttribute('aria-label'), 'White circle on a colored report cover');
    check('default failure text', await preview.locator('[data-card-image-fallback]').textContent(), 'White circle on a colored report cover');
    await target('custom-fallback').check();
    check('custom fallback fills the same viewport as the loaded image', await mediaGeometry(), loadedGeometry);
    check('custom failure slot', await target('fallback-slot').textContent(), 'Custom unavailable cover');
    await target('decorative').check();
    check('decorative failure hidden', await preview.locator('[data-card-image-fallback]').getAttribute('aria-hidden'), 'true');
    await target('logo-callback').uncheck();
    for (const horizontal of [false, true]) for (const narrow of [false, true]) for (const rtl of [false, true]) {
      await target('orientation').setChecked(horizontal);
      await target('width').setChecked(narrow);
      await panel.locator('[data-cards-demo="direction"]').setChecked(rtl);
      await target('decorative').uncheck();
      await target('custom-fallback').uncheck();
      await target('new-src').click();
      await page.waitForFunction(() => { const image = document.querySelector('[data-card-instance="regression"] [data-card-preview] > img'); return image && image.complete && image.naturalWidth > 0; });
      await settle();
      const expectedGeometry = await mediaGeometry();
      const label = `${horizontal ? 'horizontal' : 'vertical'} ${narrow ? '240' : '520'} ${rtl ? 'RTL' : 'LTR'}`;
      const screenshotCase = horizontal && !narrow && !rtl;
      if (screenshotCase) await preview.screenshot({ path: path.join(evidence, `${boundary}-media-loaded.png`) });
      await target('fail-image').click();
      await preview.locator('[data-card-image-fallback]').waitFor();
      await settle();
      check(`${label} default fallback preserves full image geometry`, await mediaGeometry(), expectedGeometry);
      if (screenshotCase) await preview.screenshot({ path: path.join(evidence, `${boundary}-media-fallback.png`) });
      await target('custom-fallback').check();
      check(`${label} custom fallback preserves full image geometry`, await mediaGeometry(), expectedGeometry);
      await target('custom-fallback').uncheck();
      await target('decorative').check();
      check(`${label} empty decorative fallback reserves the image area`, await mediaGeometry(), expectedGeometry);
    }
    await target('orientation').check();
    await target('width').check();
    await panel.locator('[data-cards-demo="direction"]').check();
    await target('new-src').click();
    check('new src clears failure', await preview.locator('[data-card-image-fallback]').count(), 0);
    await page.waitForFunction(() => { const image = document.querySelector('[data-card-instance="regression"] [data-card-preview] .fui-Image'); return image && image.complete && image.naturalWidth > 0; });
    await target('loading').check();
    check('external AutoSkeleton loading', await target('skeleton').getAttribute('aria-busy'), 'true');
    await target('skeleton').locator('[data-auto-skeleton-shape]').first().waitFor();
    check('AutoSkeleton rectangles retain rounded Fluent corners', await target('skeleton').locator('[data-auto-skeleton-shape="rectangle"]').evaluateAll(nodes => nodes.length > 0 && nodes.every(node => {
      const style = getComputedStyle(node);
      return [style.borderTopLeftRadius, style.borderTopRightRadius, style.borderBottomLeftRadius, style.borderBottomRightRadius].every(radius => parseFloat(radius) > 0);
    })), true);
    check('AutoSkeleton exposes only circle and rectangle shapes', await target('skeleton').locator('[data-auto-skeleton-shape]').evaluateAll(nodes => nodes.every(node => ['circle', 'rectangle'].includes(node.getAttribute('data-auto-skeleton-shape')))), true);
    await target('skeleton').screenshot({ path: path.join(evidence, `${boundary}-skeleton-rounded.png`) });
    await target('loading').uncheck();
    check('loading composition retains caller node', await input.evaluate(node => window.cardsRetainedInput === node && node.value === 'Retained caller edit'), true);
    await page.emulateMedia({ forcedColors: 'active' });
    check('forced colors heading', await preview.locator('h2').isVisible(), true);
    check('forced colors actions', await target('footer-action').isVisible(), true);
    await preview.screenshot({ path: path.join(evidence, `${boundary}-forced-colors.png`) });
    await page.emulateMedia({ forcedColors: 'none' });
    await page.evaluate(() => window.demoFixtureHost.update({ theme: 'light' })); await panel.locator('[data-cards-demo="direction"]').uncheck();
    await target('width').uncheck(); await target('orientation').uncheck(); await settle();
    await main.screenshot({ path: path.join(evidence, `${boundary}-light.png`) });
    await target('title-input').fill('Edited report title');
    await target('description-input').fill('Edited report description');
    check('editable title applies to actual PreviewCard', await preview.locator('h2').textContent(), 'Edited report title');
    check('editable description applies to actual PreviewCard', await preview.getByText('Edited report description', { exact: true }).count(), 1);
    const independentInput = independent.locator('[data-cards-demo=retained-input]');
    await independentInput.fill('Independent retained edit');
    await page.getByRole('button', { name: 'Independent RTL instance', exact: true }).click();
    check('collapsed independent instance inaccessible', await independentInput.evaluate(node => node.closest('.fui-AccordionPanel').hasAttribute('inert')), true);
    await page.getByRole('button', { name: 'Independent RTL instance', exact: true }).click();
    check('independent edit retained after collapse', await independentInput.inputValue(), 'Independent retained edit');
    check('main edited props isolated', await target('title-input').inputValue(), 'Edited report title');
    check('console errors', errors, []); check('console warnings', warnings, []);
    return { boundary, url, browser: browser.version(), checks, errors, warnings };
  } finally { await browser?.close(); await new Promise(resolve => server.close(resolve)); }
}
(async () => {
  fs.mkdirSync(evidence, { recursive: true });
  const reports = [];
  for (const boundary of ['source', 'compiled']) { const report = await verify(boundary); reports.push(report); console.log(`Cards ${boundary}: ${report.checks.length} browser checks passed (${report.browser}).`); }
  fs.writeFileSync(path.join(evidence, 'browser-results.json'), JSON.stringify({ browserPlugin: 'not available', fallback: 'Existing Playwright and installed Chrome; no dependency installation', sharePointHost: 'NOT EXECUTED', reports }, undefined, 2));
})().catch(error => { console.error(error); process.exitCode = 1; });
