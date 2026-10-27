// Private browser acceptance shared by source and compiled routes. Browser plugin
// not available: existing Playwright library + installed Chrome, no new dependency.
// Formatting uses DOM Selection, mouse/keyboard and Fluent controls, never .editor.
const fs = require('node:fs');
const path = require('node:path');

async function checkWebPartShellBrowser(page, record) {
  const { url, output, label, check } = record;
  const root = page.locator('[data-shell-root="first"]');
  const other = page.locator('[data-shell-root="second"]');
  const editable = root.locator('[contenteditable="true"]');
  const toolbar = () => page.locator('[data-title-toolbar]');
  const observations = () => page.evaluate(() => window.shellFixture.observe());
  const render = async patch => {
    await page.evaluate(value => window.shellFixture.render(value), patch);
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  };
  const capture = async name => {
    await page.waitForTimeout(200); // Allow Fluent's normal popup motion to finish before visual evidence.
    await page.screenshot({ path: path.join(output, `${label}-shell-${name}.png`), fullPage: false });
    fs.writeFileSync(path.join(output, `${label}-shell-${name}.html`), await page.content());
  };
  const savedHTML = async () => (await observations()).properties.titleHTML;
  const waitHTML = async fragment => { try { await page.waitForFunction(value => (window.shellFixture.observe().properties.titleHTML || '').includes(value), fragment); } catch (error) { await capture('failed-interaction'); fs.writeFileSync(path.join(output, `${label}-shell-failed-state.json`), JSON.stringify({ expected: fragment, actual: await observations(), selection: await page.evaluate(() => document.getSelection().toString()) }, null, 2)); throw error; } };
  // Uses the actual document Selection API, then waits for ProseMirror's native
  // selectionchange handler. This is DOM-driven selection, not engine commands.
  const selectRange = async (start, end) => {
    await editable.focus();
    await editable.evaluate((element, offsets) => {
      const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
      const nodes = []; let text; let total = 0;
      while ((text = walker.nextNode())) { nodes.push({ text, start: total }); total += text.textContent.length; }
      const point = offset => { const item = [...nodes].reverse().find(node => offset >= node.start) || nodes[0]; return [item.text, Math.min(offset - item.start, item.text.textContent.length)]; };
      const range = document.createRange(); const from = point(offsets[0]), to = point(offsets[1]);
      range.setStart(...from); range.setEnd(...to); const selection = document.getSelection(); selection.removeAllRanges(); selection.addRange(range);
    }, [start, end]);
    await page.waitForTimeout(60);
  };
  const focusProof = async name => { await page.waitForTimeout(100); return check(`${name}: command restores actual editor focus`, await editable.evaluate(e => e === document.activeElement), true); };
  const reset = async (title = 'Alpha Beta Gamma', titleHTML = '<h2>Alpha Beta Gamma</h2>') => {
    await render({ properties: { title, titleHTML, showTitle: true }, width: 900, surface: 'section', padding: 'auto' });
    await editable.waitFor(); await editable.focus(); await page.waitForTimeout(60);
  };
  const button = name => toolbar().getByRole('button', { name, exact: true });
  const command = async name => {
    if (await button(name).isVisible()) await button(name).click();
    else { await button('More text formatting').click(); await page.getByRole('menuitem', { name, exact: true }).or(page.getByRole('menuitemcheckbox', { name, exact: true })).click(); }
  };
  const toolbarGeometry = [];
  const geometry = async () => toolbar().evaluate(e => {
    const rect = e.getBoundingClientRect();
    const row = e.querySelector('[role="toolbar"]');
    const controls = [...row.children].filter(node => {
      const css = getComputedStyle(node), bounds = node.getBoundingClientRect();
      return node.getAttribute('aria-hidden') !== 'true' && css.display !== 'none' && bounds.width > 0;
    });
    const bounds = controls.map(node => node.getBoundingClientRect());
    const heading = e.querySelector('.fui-Dropdown').getBoundingClientRect();
    const size = e.querySelector('.fui-Combobox').getBoundingClientRect();
    const css = getComputedStyle(e), rtl = css.direction === 'rtl';
    return { width: rect.width, left: rect.left, right: rect.right, viewport: innerWidth,
      targetWidth: document.querySelector('[data-shell-root="first"] [contenteditable]').getBoundingClientRect().width,
      contentWidth: Math.max(...bounds.map(item => item.right)) - Math.min(...bounds.map(item => item.left)),
      chrome: ['paddingLeft', 'paddingRight', 'borderLeftWidth', 'borderRightWidth'].reduce((sum, property) => sum + parseFloat(css[property] || '0'), 0),
      gap: heading.width > 0 ? (rtl ? heading.left - size.right : size.left - heading.right) : undefined,
      gapToken: parseFloat(css.getPropertyValue('--spacingHorizontalS')),
      overflow: Boolean(e.querySelector('[data-overflow-menu]')),
      hidden: [...e.querySelectorAll('[data-overflowing]')].map(node => node.textContent + node.getAttribute('aria-label')),
      visible: controls.map(node => node.getAttribute('aria-label') || node.textContent),
    };
  });
  const requests = [];
  page.on('request', request => { if (/\.js(?:\?|$)/.test(request.url())) requests.push(request.url()); });
  await page.goto(url); await root.locator('h2').waitFor();
  await check('Shell page identity', await page.title(), 'WebPartShell browser fixture');
  await check('Shell local origin', new URL(page.url()).origin, url);
  await check('Shell meaningful sample', await root.locator('[data-shell-demo="panel"]').count(), 1);
  await check('Shell no framework overlay', await page.locator('webpack-dev-server-client-overlay,#webpack-dev-server-client-overlay,nextjs-portal').count(), 0);
  await check('Read requests zero editor chunks', requests.filter(request => !request.endsWith('/fixture.js')), []);
  await check('Read creates no editable engine DOM', await root.locator('[contenteditable]').count(), 0);
  await check('Read mount never writes properties', (await observations()).refreshes, [0, 0]);
  await check('metadata search/plain fixture property retained', (await observations()).properties.listId, 'keep-0');
  const headerGap = async shell => shell.locator('[data-web-part-shell-body]').evaluate(body => {
    const header = body.previousElementSibling;
    return header ? body.getBoundingClientRect().top - header.getBoundingClientRect().bottom : undefined;
  });
  await check('Read header/body default gap', await headerGap(root), 12);
  await render({ otherProperties: { showTitle: false } });
  await check('Hidden title without actions removes empty header', await other.locator('[data-web-part-shell-body]').evaluate(body => body.previousElementSibling === null), true);
  await check('Hidden title without actions leaves no blank gap', await other.locator('[data-web-part-shell-body]').evaluate(body => body.getBoundingClientRect().top - body.parentElement.getBoundingClientRect().top), 0);
  await render({ otherProperties: { showTitle: true } });
  await check('Restored title restores the default gap', await headerGap(other), 12);
  await capture('read');
  // Compare actual pixels with the published Fluent recipes in both branches.
  // Explicit sizes/line spacing are overrides, never inferred persistence.
  const typography = async element => element.evaluate(node => {
    const css = getComputedStyle(node);
    return [css.fontSize, css.lineHeight, css.fontWeight];
  });
  for (const [tag, size, line, weight] of [
    ['h1', '32px', '40px', '600'], ['h2', '28px', '36px', '600'],
    ['h3', '24px', '32px', '600'], ['h4', '20px', '28px', '600'], ['p', '14px', '20px', '400'],
  ]) {
    const html = `<${tag}>Preset title</${tag}>`;
    await render({ mode: 'Read', properties: { title: 'Preset title', titleHTML: html } });
    await check(`${tag}: Read uses exact Fluent typography`, await typography(root.locator(tag).filter({ hasText: /^Preset title$/ })), [size, line, weight]);
    await page.getByRole('button', { name: 'Enter page Edit mode', exact: true }).click(); await editable.waitFor(); await editable.focus();
    await check(`${tag}: Edit retains header/body gap`, await headerGap(root), 12);
    await check(`${tag}: Edit uses identical Fluent typography`, await typography(editable.locator(tag)), [size, line, weight]);
    await check(`${tag}: picker shows inherited preset size`, await toolbar().getByRole('combobox', { name: 'Font size', exact: true }).inputValue(), size.slice(0, -2));
    await check(`${tag}: mounting Read/Edit does not persist inherited styles`, await savedHTML(), html);
  }
  await render({ mode: 'Read', fluentTheme: { fontSizeHero700: '40px', lineHeightHero700: '48px' }, properties: { titleHTML: '<h2>Theme title</h2>' } });
  await check('Custom theme: Read resolves typography tokens', await typography(root.locator('h2')), ['40px', '48px', '600']);
  await page.getByRole('button', { name: 'Enter page Edit mode', exact: true }).click(); await editable.waitFor(); await editable.focus();
  await check('Custom theme: Edit resolves the same tokens', await typography(editable.locator('h2')), ['40px', '48px', '600']);
  await check('Custom theme: picker matches rendered size', await toolbar().getByRole('combobox', { name: 'Font size', exact: true }).inputValue(), '40');
  await render({ fluentTheme: { fontSizeHero700: '44px', lineHeightHero700: '52px' } });
  await check('Live theme: picker updates without editing title', await toolbar().getByRole('combobox', { name: 'Font size', exact: true }).inputValue(), '44');
  await check('Live theme: inherited tokens never become HTML overrides', await savedHTML(), '<h2>Theme title</h2>');
  await render({ properties: { titleHTML: '<h2 class="wps-line-20"><span class="wps-size-20">Explicit</span></h2>' } }); await editable.focus();
  await check('Explicit line spacing wins over preset token', await editable.locator('h2').evaluate(e => getComputedStyle(e).lineHeight), '88px');
  await check('Explicit size wins over theme token', await editable.locator('span').evaluate(e => getComputedStyle(e).fontSize), '20px');
  await button('Spacing').click(); await page.getByRole('menuitemradio', { name: 'Default line spacing', exact: true }).click();
  await check('Default spacing restores theme token in Edit', await editable.locator('h2').evaluate(e => getComputedStyle(e).lineHeight), '52px');
  await check('Default spacing retains explicit size in HTML', await savedHTML(), '<h2><span class="wps-size-20">Explicit</span></h2>');
  await render({ mode: 'Read' });
  await check('Default spacing has identical Read output', await root.locator('h2').evaluate(e => getComputedStyle(e).lineHeight), '52px');
  await render({ mode: 'Read', fluentTheme: undefined, properties: { title: 'Alpha Beta Gamma', titleHTML: undefined } });
  // Fresh host theme reproduces the same-mask interval from the independent
  // review: budget changes alone must refit the currently visible command row.
  for (const dir of ['ltr', 'rtl']) {
    await page.setViewportSize({ width: 320, height: 900 });
    await render({ mode: 'Read', width: 320, dir });
    await page.getByRole('button', { name: 'Enter page Edit mode', exact: true }).click();
    await editable.focus(); await toolbar().waitFor();
    let narrowMask;
    for (const size of [320, 321, 325, 320, 900, 325, 900, 320, 900, 325, 321, 320]) {
      await page.setViewportSize({ width: size, height: 900 }); await render({ width: size });
      await editable.focus(); await page.waitForTimeout(250); const first = await geometry();
      await page.waitForTimeout(450); const settled = await geometry();
      toolbarGeometry.push({ cycle: 'same-mask', dir, size, first, settled });
      await check(`${dir}/${size}: same-mask budget change fits visible controls plus chrome`, Math.abs(settled.width - settled.contentWidth - settled.chrome) <= 1, true);
      await check(`${dir}/${size}: same-mask budget change settles width/mask`, first, settled);
      await check(`${dir}/${size}: same-mask actual combo gap matches theme token`, settled.gap >= settled.gapToken - 0.1, true);
      if (size === 900) await check(`${dir}: cross-mask widening recovers every command`, settled.overflow, false);
      else {
        if (narrowMask === undefined) narrowMask = settled.hidden;
        await check(`${dir}/${size}: regression retains the same real narrow command mask`, settled.hidden, narrowMask);
      }
    }
    await capture(`same-mask-${dir}`);
  }
  await page.setViewportSize({ width: 1210, height: 900 }); await render({ mode: 'Read', width: 900, dir: 'ltr' });

  await root.locator('[data-shell-demo="note"]').fill('Persistent browser note');
  await root.locator('[data-shell-demo="increment"]').click();

  for (const host of ['SharePoint', 'SingleWebPartAppPage', 'Teams', 'Office', 'Outlook']) {
    await render({ host, surface: 'auto' });
    const surface = host === 'SharePoint' ? 'section' : 'application';
    await check(`${host}: real host detection`, await root.getAttribute('data-web-part-shell-surface'), surface);
  }
  await render({ host: 'Teams', surface: 'application' });
  for (const size of [320, 500, 639, 640, 900]) {
    await render({ width: size });
    await check(`${size}: root is measured container`, await root.evaluate(e => e.getBoundingClientRect().width), size);
    await check(`${size}: exact contextual padding`, await root.locator('[data-web-part-shell-content]').evaluate(e => {
      const css = getComputedStyle(e); return [css.paddingBlockStart, css.paddingInlineStart, css.paddingInlineEnd];
    }), size < 640 ? ['12px', '12px', '12px'] : ['16px', '16px', '16px']);
    await check(`${size}: independent region keeps medium padding`, await other.locator('[data-web-part-shell-content]').evaluate(e => getComputedStyle(e).paddingInlineStart), '12px');
    await capture(`width-${size}`);
  }
  await render({ surface: 'section', width: 900 });
  await check('Section has exact zero padding', await root.locator('[data-web-part-shell-content]').evaluate(e => getComputedStyle(e).paddingInlineStart), '0px');
  for (const [padding, expected] of [['none', '0px'], ['small', '8px'], ['medium', '12px'], ['large', '16px']]) {
    await render({ padding, surface: 'application' });
    await check(`Explicit ${padding} has no responsive override`, await root.locator('[data-web-part-shell-content]').evaluate(e => getComputedStyle(e).paddingInlineStart), expected);
  }
  await render({ padding: 'auto', surface: 'section' });
  const themes = {};
  for (const theme of ['light', 'dark', 'highContrast']) {
    await render({ theme, dir: 'rtl' });
    themes[theme] = await root.evaluate(e => ({ foreground: getComputedStyle(e).getPropertyValue('--colorNeutralForeground1').trim(), direction: getComputedStyle(e).direction }));
    await check(`${theme}: actual RTL layout`, themes[theme].direction, 'rtl');
    await capture(`${theme}-rtl`);
  }
  await check('Live SDK theme hooks dark foreground', themes.dark.foreground, '#ffffff');
  await check('Live SDK theme hooks contrast foreground', themes.highContrast.foreground, '#ffffff');
  await render({ dir: 'ltr', theme: 'light' });
  await page.emulateMedia({ forcedColors: 'active' });
  await check('Forced colors media actually active', await page.evaluate(() => matchMedia('(forced-colors: active)').matches), true);
  await capture('forced-colors'); await page.emulateMedia({ forcedColors: 'none' });

  await page.getByRole('button', { name: 'Enter page Edit mode', exact: true }).click(); await editable.waitFor();
  await check('Edit requests actual async engine assets', requests.some(request => request.includes('editable-text-editor')), true);
  // Begin narrow before the first toolbar mount, then repeatedly recover the full
  // row. Hidden items must never redefine the intrinsic full-content footprint.
  await page.setViewportSize({ width: 320, height: 900 }); await render({ width: 320 });
  await editable.click(); await toolbar().waitFor();
  for (const [cycle, dir, sizes] of [[0, 'ltr', [320, 500, 639, 640, 900]], [1, 'rtl', [320, 500, 639, 640, 900]], [2, 'ltr', [900, 320, 900]]]) {
    for (const size of sizes) {
      await page.setViewportSize({ width: size, height: 900 }); await render({ width: size, dir });
      await editable.focus(); await page.waitForTimeout(250);
      const first = await geometry(); await page.waitForTimeout(450); const settled = await geometry();
      toolbarGeometry.push({ cycle, dir, size, first, settled });
      await check(`${cycle}/${dir}/${size}: toolbar width and overflow state settle without oscillation`, first, settled);
      await check(`${cycle}/${dir}/${size}: toolbar surface stays target-safe and within viewport`, settled.width <= settled.targetWidth + 1 && settled.left >= -1 && settled.right <= settled.viewport + 1, true);
      await check(`${cycle}/${dir}/${size}: toolbar has no clipped overflow`, await toolbar().evaluate(e => e.scrollWidth <= e.clientWidth + 1), true);
      await check(`${cycle}/${dir}/${size}: surface fits the actual visible controls plus its padding/border`, Math.abs(settled.width - settled.contentWidth - settled.chrome) <= 1, true);
      if (settled.gap !== undefined) await check(`${cycle}/${dir}/${size}: complete combo controls have the theme token gap`, settled.gap >= settled.gapToken - 0.1, true);
      // The medium icon buttons need more room than the former small buttons
      // row. 640/900 must recover fully; at 500 every hidden command must instead
      // be reachable through Fluent's actual overflow menu.
      if (size >= 640) await check(`${cycle}/${dir}/${size}: widening recovers every command`, settled.overflow, false);
      if (size === 500 && settled.overflow) {
        await button('More text formatting').click();
        for (const name of ['Heading', 'Bold', 'Italic', 'Underline', 'Theme color', 'Clear formatting', 'Spacing']) {
          const direct = toolbar().getByRole(name === 'Heading' ? 'combobox' : 'button', { name, exact: true });
          const menu = page.getByRole('menuitem', { name, exact: true }).or(page.getByRole('menuitemcheckbox', { name, exact: true }));
          await check(`${cycle}/${dir}/${size}: ${name} remains reachable`, await direct.isVisible() || await menu.isVisible(), true);
        }
        await page.keyboard.press('Escape');
      }
      await check(`${cycle}/${dir}/${size}: V1 omits painter`, await toolbar().getByRole('button', { name: 'Format painter', exact: true }).count(), 0);
      await check(`${cycle}/${dir}/${size}: font control replaces Fluent min-width`, await toolbar().locator('.fui-Combobox').evaluate(e => getComputedStyle(e).minWidth), '64px');
      await page.keyboard.press('Alt+F10');
      await check(`${cycle}/${dir}/${size}: overflow keyboard entry focuses visible control`, await page.evaluate(() => Boolean(document.activeElement.closest('[data-title-toolbar]')) && !document.activeElement.closest('[data-overflowing]')), true);
      const focused = await page.evaluate(() => document.activeElement.getAttribute('aria-label'));
      await page.keyboard.press('Tab');
      await check(`${cycle}/${dir}/${size}: Tab moves from toolbar control`, await page.evaluate(before => document.activeElement.getAttribute('aria-label') !== before, focused), true);
      await editable.focus();
      if (cycle === 0) await capture(`edit-width-${size}`);
      await page.locator('[data-fixture-outside]').click(); await toolbar().waitFor({ state: 'detached' });
      await editable.focus(); await toolbar().waitFor(); await page.waitForTimeout(250);
      const reopened = await geometry();
      await check(`${cycle}/${dir}/${size}: hide/show retains the settled footprint and overflow commands`, Math.abs(reopened.width - settled.width) <= 1 && JSON.stringify(reopened.hidden) === JSON.stringify(settled.hidden), true);
    }
  }
  await page.setViewportSize({ width: 320, height: 900 });
  await render({ mode: 'Read', width: 320, dir: 'rtl' });
  await render({ mode: 'Edit' }); await editable.focus(); await toolbar().waitFor();
  await page.waitForTimeout(250); const freshRtl = await geometry();
  await page.waitForTimeout(450); const freshRtlSettled = await geometry();
  toolbarGeometry.push({ cycle: 'fresh-edit', dir: 'rtl', size: 320, first: freshRtl, settled: freshRtlSettled });
  await check('First Edit mount at narrow RTL: visible content fits without unused space', Math.abs(freshRtlSettled.width - freshRtlSettled.contentWidth - freshRtlSettled.chrome) <= 1, true);
  await check('First Edit mount at narrow RTL: width/mask settle without oscillation', freshRtl, freshRtlSettled);
  await check('First Edit mount at narrow RTL: actual combo gap matches token', freshRtlSettled.gap >= freshRtlSettled.gapToken - 0.1, true);
  await capture('edit-first-narrow-rtl');
  // A viewport resize can change the available budget without resizing the
  // editor target. It must also restore full commands after narrowing.
  await page.setViewportSize({ width: 900, height: 900 }); await render({ width: 900, dir: 'ltr' });
  for (const size of [320, 900]) {
    await page.setViewportSize({ width: size, height: 900 }); await editable.focus();
    await page.waitForTimeout(250); const first = await geometry();
    await page.waitForTimeout(450); const settled = await geometry();
    toolbarGeometry.push({ cycle: 'viewport-only', dir: 'ltr', size, first, settled });
    await check(`${size}: viewport-only resize fits visible content plus chrome`, Math.abs(settled.width - settled.contentWidth - settled.chrome) <= 1, true);
    await check(`${size}: viewport-only resize settles width/mask`, first, settled);
    await check(`${size}: viewport-only resize stays within viewport`, settled.left >= -1 && settled.right <= settled.viewport + 1, true);
    if (size === 900) await check('Viewport-only widening recovers all formatting controls', settled.overflow, false);
  }
  fs.writeFileSync(path.join(output, `${label}-shell-toolbar-geometry.json`), JSON.stringify(toolbarGeometry, null, 2));
  await page.setViewportSize({ width: 320, height: 900 }); await render({ width: 320, dir: 'rtl' });
  await selectRange(0, 5); await button('More text formatting').click();
  await check('Narrow RTL overflow excludes painter', await page.getByRole('menuitemcheckbox', { name: 'Format painter', exact: true }).count(), 0);
  // The larger preset can wrap the header action, leaving Italic visible.
  // Underline remains overflowed in this native layout and exercises the menu.
  await page.getByRole('menuitemcheckbox', { name: 'Underline', exact: true }).click();
  await waitHTML('<u>'); await focusProof('Narrow RTL overflow underline');
  await check('Narrow RTL overflow command preserves native range length', await page.evaluate(() => document.getSelection().toString().length), 5);
  await render({ dir: 'ltr' });
  await page.setViewportSize({ width: 1210, height: 900 }); await render({ width: 900, properties: { title: "We're here to help", titleHTML: "<h2>We're here to help</h2>" } }); await editable.focus(); await capture('native-comparison'); await reset();
  await check('Edit retains child note', await root.locator('[data-shell-demo="note"]').inputValue(), 'Persistent browser note');
  await check('Edit retains child count', (await root.locator('[data-shell-demo="observations"]').textContent()).includes('"count":1'), true);
  await check('No child remount across host/theme/mode changes', (await observations()).mounts, 1);
  await selectRange(0, 5); await command('Bold'); await waitHTML('<strong>Alpha</strong>'); await focusProof('Bold');
  await command('Italic'); await waitHTML('<em>'); await focusProof('Italic');
  await command('Underline'); await waitHTML('<u>'); await focusProof('Underline');
  await check('Native selection remains original range after repeated marks', await page.evaluate(() => document.getSelection().toString()), 'Alpha');
  await toolbar().getByRole('combobox', { name: 'Font size', exact: true }).fill('37');
  await page.keyboard.press('Enter'); await waitHTML('wps-size-37'); await focusProof('Size');
  await toolbar().getByRole('combobox', { name: 'Font size', exact: true }).click();
  await page.getByRole('option', { name: '28', exact: true }).click(); await waitHTML('wps-size-28'); await focusProof('Size option');
  await toolbar().getByRole('combobox', { name: 'Heading', exact: true }).click();
  await page.getByRole('option', { name: 'Heading 3', exact: true }).click(); await waitHTML('<h3>'); await focusProof('Heading');
  for (const [choice, block] of [['Heading 1', 'h1'], ['Heading 2', 'h2'], ['Heading 4', 'h4'], ['Normal', 'p'], ['Heading 3', 'h3']]) {
    await toolbar().getByRole('combobox', { name: 'Heading', exact: true }).click();
    await page.getByRole('option', { name: choice, exact: true }).click(); await waitHTML(`<${block}>`); await focusProof(choice);
  }
  const beforeInvalidSize = await savedHTML(); const beforeInvalidRefresh = (await observations()).refreshes[0];
  await toolbar().getByRole('combobox', { name: 'Font size', exact: true }).fill('20.5'); await page.keyboard.press('Enter');
  await focusProof('Invalid size'); await check('Invalid fractional font input never writes', await savedHTML(), beforeInvalidSize);
  await check('Invalid fractional font has no property refresh', (await observations()).refreshes[0], beforeInvalidRefresh);
  for (const color of ['brand', 'subtle', 'primary']) {
    await button('Theme color').click();
    await page.locator(`[data-title-color=${color}]`).click();
    await waitHTML(`wps-color-${color.toLowerCase()}`); await focusProof(`Palette ${color}`);
    await check(`Palette ${color}: parent toolbar continuously visible`, await toolbar().isVisible(), true);
  }
  await button('Spacing').click(); await check('Spacing exposes default and six finite overrides', await page.getByRole('menuitemradio').count(), 7); await page.getByRole('menuitemradio', { name: 'Line spacing 2', exact: true }).click(); await waitHTML('wps-line-20');
  await button('Spacing').click(); await page.getByRole('menuitemcheckbox', { name: 'Space before', exact: true }).click(); await waitHTML('wps-space-before');
  await button('Spacing').click(); await page.getByRole('menuitemcheckbox', { name: 'Space after', exact: true }).click(); await waitHTML('wps-space-after'); await focusProof('Spacing');
  await capture('formatted-toolbar');
  await editable.evaluate(e => { window.__shellThemeEditable = e; });
  const beforeThemeRefresh = (await observations()).refreshes[0];
  for (const theme of ['light', 'dark', 'highContrast']) {
    await render({ theme, dir: 'rtl' });
    await check(`${theme}: live Edit theme retains editor DOM`, await editable.evaluate(e => e === window.__shellThemeEditable), true);
    await check(`${theme}: live Edit theme retains native range`, await page.evaluate(() => document.getSelection().toString()), 'Alpha');
    await capture(`editor-${theme}-rtl`);
  }
  await page.emulateMedia({ forcedColors: 'active' }); await capture('editor-forced-colors'); await page.emulateMedia({ forcedColors: 'none' });
  await render({ theme: 'light', dir: 'ltr' });
  await check('Theme/RTL changes do not publish title properties', (await observations()).refreshes[0], beforeThemeRefresh);
  const beforeClear = await savedHTML();
  await command('Clear formatting'); await waitHTML('<h3>Alpha Beta Gamma</h3>');
  await page.keyboard.press('Meta+z'); await page.waitForFunction(value => window.shellFixture.observe().properties.titleHTML === value, beforeClear);
  await check('Clear formatting is one native undo', await savedHTML(), beforeClear);

  await reset('Alpha Beta Gamma', '<h2><strong>Alpha</strong> Beta Gamma</h2>');
  await check('V1 toolbar has no Format painter', await toolbar().getByRole('button', { name: 'Format painter', exact: true }).count(), 0);
  await selectRange(6, 6); await page.keyboard.down('Shift');
  for (let index = 0; index < 4; index++) await page.keyboard.press('ArrowRight');
  await page.keyboard.up('Shift'); await page.waitForTimeout(100);
  await check('Native Shift selection does not copy formatting', await savedHTML(), '<h2><strong>Alpha</strong> Beta Gamma</h2>');
  await command('Italic'); await waitHTML('<em>Beta</em>'); await focusProof('Native Shift range italic');
  await check('Native Shift selection remains target after command', await page.evaluate(() => document.getSelection().toString()), 'Beta');
  await page.keyboard.press('Meta+z'); await page.waitForFunction(() => !(window.shellFixture.observe().properties.titleHTML || '').includes('<em>Beta</em>'));
  await check('Native range command is one undo', await savedHTML(), '<h2><strong>Alpha</strong> Beta Gamma</h2>');
  await selectRange(6, 6);
  const beforeWholeRefresh = (await observations()).refreshes[0];
  await command('Bold');
  await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML === '<h2><strong>Alpha Beta Gamma</strong></h2>');
  await check('No selection makes mixed Bold uniform across the whole title', await savedHTML(), '<h2><strong>Alpha Beta Gamma</strong></h2>');
  await check('Whole-title formatting publishes once', (await observations()).refreshes[0], beforeWholeRefresh + 1);
  await check('Whole-title command preserves the native caret offset', await editable.evaluate(element => {
    const selection = element.ownerDocument.getSelection();
    if (!selection?.isCollapsed || !selection.anchorNode || !element.contains(selection.anchorNode)) return -1;
    const prefix = element.ownerDocument.createRange(); prefix.selectNodeContents(element);
    prefix.setEnd(selection.anchorNode, selection.anchorOffset); return prefix.toString().length;
  }), 6);
  await command('Bold');
  await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML === '<h2>Alpha Beta Gamma</h2>');
  await check('Second caret Bold removes emphasis from the whole title', await savedHTML(), '<h2>Alpha Beta Gamma</h2>');
  await page.keyboard.press('Meta+z');
  await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML === '<h2><strong>Alpha Beta Gamma</strong></h2>');
  await check('Whole-title toggle is one native undo', await savedHTML(), '<h2><strong>Alpha Beta Gamma</strong></h2>');
  await page.keyboard.press('Meta+b');
  await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML === '<h2>Alpha Beta Gamma</h2>');
  await check('Caret keyboard Bold follows the same whole-title scope', await savedHTML(), '<h2>Alpha Beta Gamma</h2>');
  await reset('Alpha Beta Gamma', '<h2><strong>Alpha</strong> Beta Gamma</h2>');
  await selectRange(6, 6); await command('Underline');
  await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML === '<h2><strong><u>Alpha</u></strong><u> Beta Gamma</u></h2>');
  await check('Caret command formats the entire title', await savedHTML(), '<h2><strong><u>Alpha</u></strong><u> Beta Gamma</u></h2>');
  await page.keyboard.insertText('X');
  await page.waitForFunction(() => window.shellFixture.observe().properties.titleHTML === '<h2><strong><u>Alpha</u></strong><u> XBeta Gamma</u></h2>');
  await check('Caret command retains its position and future input formatting', await savedHTML(), '<h2><strong><u>Alpha</u></strong><u> XBeta Gamma</u></h2>');
  await other.locator('[contenteditable]').click();
  await check('Other editor dismisses toolbar without formatting its document', (await observations()).otherProperties.titleHTML, undefined);
  await page.locator('[data-fixture-outside]').click();
  await check('Genuine outside click dismisses all title toolbars', await toolbar().count(), 0);
  await reset(); await editable.click(); await page.keyboard.press('Alt+F10');
  await check('Alt+F10 focuses actual visible toolbar control', await page.evaluate(() => Boolean(document.activeElement.closest('[data-title-toolbar]'))), true);
  await page.keyboard.press('Tab');
  await check('Tab remains a real focus movement', await page.evaluate(() => document.activeElement.tagName !== 'BODY'), true);
  await page.keyboard.press('Escape');
  await editable.focus(); await page.keyboard.press('Escape');
  await check('Escape hides toolbar', await toolbar().count(), 0);
  await editable.focus(); await page.keyboard.press('Meta+a'); await page.keyboard.press('Meta+b'); await page.keyboard.insertText('Keyboard title'); await waitHTML('<strong>Keyboard title</strong>');
  await check('Keyboard shortcut publishes atomic plain and rich title', (await observations()).properties.title, 'Keyboard title');
  await check('First editor does not change second properties', (await observations()).otherProperties.title, 'Independent title');

  await editable.evaluate(e => { window.__shellEditable = e; });
  await page.keyboard.press('Meta+a');
  await render({ properties: { title: 'Authoritative title', titleHTML: '<p>Authoritative title</p>' } });
  await page.waitForFunction(() => document.querySelector('[data-shell-root="first"] [contenteditable]').textContent === 'Authoritative title');
  await check('External replacement retains same editor DOM', await editable.evaluate(e => e === window.__shellEditable), true);
  await page.keyboard.press('Meta+z');
  await check('External replacement clears native undo history', await editable.textContent(), 'Authoritative title');
  await page.keyboard.insertText('New');
  await page.keyboard.press('Meta+z');
  await check('New edit undo returns authoritative base', await editable.textContent(), 'Authoritative title');
  await check('External SelectAll conversion retains authoritative paragraph', await editable.locator('p').count(), 1);

  await reset('Paste base', '<h2>Paste base</h2>'); await page.keyboard.press('Meta+a');
  // Synthetic ClipboardEvent is explicit: native DataTransfer/codec path, not
  // operating-system clipboard permission or browser menu paste.
  await editable.evaluate(e => { const data = new DataTransfer(); data.setData('text/html', '<h3><strong>Pasted</strong></h3><p>Second <script>bad()</script>line</p>'); data.setData('text/plain', 'Pasted\nSecond line'); e.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data })); });
  await waitHTML('<strong>Pasted</strong>');
  await check('DOM clipboard paste sanitizes script', !(await savedHTML()).includes('script') && !(await savedHTML()).includes('bad()'), true);
  await page.keyboard.press('ArrowRight'); await page.keyboard.press('Enter'); await page.keyboard.insertText('End'); await waitHTML('<br>');
  await check('Native Enter keeps one title block', await editable.locator('h1,h2,h3,h4,p').count(), 1);
  await capture('paste-enter');

  await reset('Composition base', '<h2>Composition base</h2>'); await page.keyboard.press('Meta+a');
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: '編集中', selectionStart: 3, selectionEnd: 3 });
  await cdp.send('Input.insertText', { text: '編集中' });
  await page.waitForFunction(() => window.shellFixture.observe().properties.title === '編集中');
  await check('Chromium CDP composition publishes final unicode title', (await observations()).properties.title, '編集中');
  await cdp.detach();

  await root.locator('[data-shell-demo="empty"]').click();
  await check('Defined empty rich title preserved', await savedHTML(), '');
  await root.locator('[data-shell-demo="action"]').click();
  await check('Hook action hides editor', await editable.count(), 0);
  await check('Hidden title removes the complete action header', await root.locator('[data-web-part-shell-header], [data-shell-demo="action"]').count(), 0);
  await check('Hidden title with actions leaves no blank gap', await root.locator('[data-web-part-shell-body]').evaluate(body => body.getBoundingClientRect().top - body.parentElement.getBoundingClientRect().top), 0);
  await root.locator('[data-shell-demo="show-title"]').check(); await editable.waitFor();
  await check('Panel can recover hidden title', await editable.count(), 1);
  await render({ mode: 'Read' });
  await check('Empty Read has no fake placeholder title', await root.locator('h1,h2,h3,h4,p').filter({ hasText: '編集中' }).count(), 0);
  await check('Read destroys editable DOM', await root.locator('[contenteditable]').count(), 0);
  await check('Child state still persistent after editor teardown', await root.locator('[data-shell-demo="note"]').inputValue(), 'Persistent browser note');
  await check('Child did not unmount', (await observations()).cleanups, 0);
  await root.locator('[data-shell-demo="plain"]').click();
  await check('Actual sample plain action renders its current subject', await root.locator('h2').textContent(), 'Project documents');
  await render({ properties: { title: 'Literal <sample> title', titleHTML: undefined } });
  await check('Plain fallback renders angle brackets as literal text', await root.locator('h2').textContent(), 'Literal <sample> title');
  await root.locator('[data-shell-demo="rich"]').click();
  await check('Actual sample rich control renders formatting', await root.locator('h2 strong').textContent(), 'Project');
  await render({ surface: undefined, padding: undefined, dir: undefined });
  await root.locator('[data-shell-demo="surface"]').selectOption('section');
  await check('Actual sample surface override affects root', await root.getAttribute('data-web-part-shell-surface'), 'section');
  await root.locator('[data-shell-demo="padding"]').selectOption('small');
  await check('Actual sample padding override affects layout', await root.locator('[data-web-part-shell-content]').evaluate(e => getComputedStyle(e).paddingInlineStart), '8px');
  await root.locator('[data-shell-demo="direction"]').selectOption('rtl');
  await check('Actual sample direction affects root', await root.evaluate(e => getComputedStyle(e).direction), 'rtl');
  await render({ theme: 'dark', host: 'Teams' });
  await check('Actual sample inherits host theme', await root.evaluate(e => getComputedStyle(e).getPropertyValue('--colorNeutralForeground1').trim()), '#ffffff');
  await render({ theme: 'light' });
  await capture('sample-controls');

  // Card is a structural boundary: subtle/Card transitions may remount the body.
  // These scenarios follow the legacy lifetime assertions and establish their
  // own baseline after the first transition. All non-subtle updates preserve it.
  await render({ host: 'SharePoint', mode: 'Read', theme: 'light', dir: 'ltr', width: 900,
    appearance: undefined, size: undefined, titlePosition: undefined, properties: { title: 'Card project heading', titleHTML: undefined, showTitle: true } });
  await root.evaluate(element => { window.__shellCardOuterRoot = element; });
  await other.getByRole('textbox', { name: 'Other child', exact: true }).fill('Independent Card state');
  const appearanceSelect = () => root.locator('[data-shell-demo="appearance"]');
  const positionSelect = () => root.locator('[data-shell-demo="title-position"]');
  const sizeSelect = () => root.locator('[data-shell-demo="card-size"]');
  const card = () => root.locator('[data-web-part-shell-card]');
  const cardHeader = () => root.locator('[data-web-part-shell-header]');
  await check('Actual sample defaults to subtle', await appearanceSelect().inputValue(), 'subtle');
  await check('Actual sample defaults Card size to medium', await sizeSelect().inputValue(), 'medium');
  await check('Actual sample defaults title outside', await positionSelect().inputValue(), 'outside');
  await check('Default subtle has no Card', await card().count(), 0);
  await positionSelect().selectOption('inside');
  await check('Subtle ignores inside placement', await cardHeader().evaluate(header => header.parentElement.hasAttribute('data-web-part-shell-content')), true);
  for (const size of ['small', 'medium', 'large']) {
    await sizeSelect().selectOption(size);
    await check(`Subtle ignores ${size} Card size`, [await card().count(), await root.getAttribute('size')], [0, null]);
  }
  await sizeSelect().selectOption('medium');
  await appearanceSelect().selectOption('filled');
  await check('Actual sample creates native Fluent Card', await card().evaluate(element => element.classList.contains('fui-Card')), true);
  await check('Options survive subtle/Card body remount', await positionSelect().inputValue(), 'inside');
  await root.locator('[data-shell-demo="note"]').fill('Card persistent note');
  await root.locator('[data-shell-demo="increment"]').click();
  const cardBaseline = await observations();
  const retainedCardState = async name => {
    await check(`${name}: body note retained`, await root.locator('[data-shell-demo="note"]').inputValue(), 'Card persistent note');
    await check(`${name}: body counter retained`, JSON.parse(await root.locator('[data-shell-demo="observations"]').textContent()).count, 1);
    await check(`${name}: body lifetime retained`, [(await observations()).mounts, (await observations()).cleanups], [cardBaseline.mounts, cardBaseline.cleanups]);
    await check(`${name}: outer root retained`, await root.evaluate(element => element === window.__shellCardOuterRoot), true);
    await check(`${name}: second instance state isolated`, await other.getByRole('textbox', { name: 'Other child', exact: true }).inputValue(), 'Independent Card state');
  };
  for (const appearance of ['filled', 'filled-alternative', 'outline']) {
    await appearanceSelect().selectOption(appearance);
    for (const position of ['outside', 'inside']) {
      await positionSelect().selectOption(position);
      const name = `${appearance}/${position}`;
      await check(`${name}: one native Card and body`, [await card().count(), await card().locator('[data-web-part-shell-body]').count()], [1, 1]);
      await check(`${name}: complete header in requested region`, await cardHeader().evaluate((header, inside) => header.parentElement.hasAttribute(inside ? 'data-web-part-shell-card' : 'data-web-part-shell-content'), position === 'inside'), true);
      const recipe = await card().evaluate((element, selectedAppearance) => {
        const css = getComputedStyle(element), after = getComputedStyle(element, '::after');
        const backgroundToken = selectedAppearance === 'filled' ? '--colorNeutralBackground1'
          : selectedAppearance === 'filled-alternative' ? '--colorNeutralBackground2' : '--colorTransparentBackground';
        const strokeToken = selectedAppearance === 'outline' ? '--colorNeutralStroke1' : '--colorTransparentStroke';
        // Canvas uses the browser's color parser without changing document styles;
        // normalize token hex/transparent and computed rgb/rgba to the same form.
        const context = document.createElement('canvas').getContext('2d');
        if (!context) throw new Error('Native Card recipe check requires a 2D color context');
        const normalize = value => {
          if (!CSS.supports('color', value)) throw new Error(`Invalid native Card color: ${value}`);
          context.fillStyle = value;
          return context.fillStyle;
        };
        return {
          background: normalize(css.backgroundColor),
          expectedBackground: normalize(css.getPropertyValue(backgroundToken).trim()),
          shadow: css.boxShadow,
          stroke: ['borderTopColor', 'borderRightColor', 'borderBottomColor', 'borderLeftColor'].map(property => normalize(after[property])),
          expectedStroke: normalize(css.getPropertyValue(strokeToken).trim()),
        };
      }, appearance);
      await check(`${name}: native appearance background recipe`, recipe.background, recipe.expectedBackground);
      await check(`${name}: native appearance shadow recipe`, recipe.shadow === 'none', appearance === 'outline');
      await check(`${name}: native appearance border recipe`, recipe.stroke, Array(4).fill(recipe.expectedStroke));
      await check(`${name}: native medium vertical spacing`, await card().evaluate(element => {
        const css = getComputedStyle(element);
        return [css.paddingTop, css.paddingRight, css.paddingBottom, css.paddingLeft, css.rowGap, css.flexDirection];
      }), ['12px', '12px', '12px', '12px', '12px', 'column']);
      await check(`${name}: existing small outer padding retained`, await root.locator('[data-web-part-shell-content]').evaluate(element => getComputedStyle(element).paddingInlineStart), '8px');
      await check(`${name}: header gap`, position === 'inside' ? await headerGap(root) : await card().evaluate(element => element.getBoundingClientRect().top - element.previousElementSibling.getBoundingClientRect().bottom), 12);
      await root.locator('[data-shell-demo="action"]').click();
      await check(`${name}: hidden removes full header/action`, await root.locator('[data-web-part-shell-header], [data-shell-demo="action"]').count(), 0);
      await check(`${name}: hidden body is first Card child without header gap`, await card().evaluate(element => {
        const css = getComputedStyle(element), body = element.querySelector('[data-web-part-shell-body]');
        return element.firstElementChild === body && Math.abs(body.getBoundingClientRect().top - element.getBoundingClientRect().top - parseFloat(css.borderTopWidth) - parseFloat(css.paddingTop)) < 1;
      }), true);
      await root.locator('[data-shell-demo="show-title"]').check();
      await check(`${name}: body switch restores full header/action`, [await cardHeader().count(), await root.locator('[data-shell-demo="action"]').count()], [1, 1]);
      await retainedCardState(name);
      await capture(`card-${appearance}-${position}`);
    }
  }
  // Each native size changes only the Card recipe. Compare identity within a
  // fixed header placement because changing placement can remount the header.
  for (const position of ['outside', 'inside']) {
    await positionSelect().selectOption(position);
    await root.evaluate(element => {
      window.__shellSizeCard = element.querySelector('[data-web-part-shell-card]');
      window.__shellSizeHeader = element.querySelector('[data-web-part-shell-header]');
      window.__shellSizeBody = element.querySelector('[data-web-part-shell-body]');
    });
    for (const [size, spacing, radius] of [['small', '8px', '2px'], ['medium', '12px', '4px'], ['large', '16px', '6px']]) {
      await sizeSelect().selectOption(size);
      const name = `Card ${position}/${size}`;
      const recipe = await card().evaluate(element => {
        const css = getComputedStyle(element);
        return [css.paddingTop, css.paddingRight, css.paddingBottom, css.paddingLeft, css.rowGap, css.borderTopLeftRadius, css.getPropertyValue('--fui-Card--size').trim()];
      });
      await check(`${name}: native size padding/gap/radius`, recipe, [spacing, spacing, spacing, spacing, spacing, radius, spacing]);
      await check(`${name}: outer inset unchanged`, await root.locator('[data-web-part-shell-content]').evaluate(element => getComputedStyle(element).paddingInlineStart), '8px');
      await check(`${name}: header spacing`, position === 'inside' ? await headerGap(root) : await card().evaluate(element => element.getBoundingClientRect().top - element.previousElementSibling.getBoundingClientRect().bottom), position === 'inside' ? parseFloat(spacing) : 12);
      await check(`${name}: Card/header/body identity retained`, await root.evaluate(element =>
        element.querySelector('[data-web-part-shell-card]') === window.__shellSizeCard &&
        element.querySelector('[data-web-part-shell-header]') === window.__shellSizeHeader &&
        element.querySelector('[data-web-part-shell-body]') === window.__shellSizeBody), true);
      await check(`${name}: size never reaches root attribute`, await root.getAttribute('size'), null);
      await retainedCardState(name);
    }
  }
  await sizeSelect().selectOption('medium');
  // Surface spacing belongs outside the native Card border. The Card owns a
  // distinct 12px inner inset in every responsive/explicit padding combination.
  for (const [width, padding, expectedInset] of [[320, 'auto', 12], [640, 'auto', 16], [900, 'small', 8]]) {
    await render({ surface: 'application', padding, width });
    for (const position of ['inside', 'outside']) {
      await positionSelect().selectOption(position);
      const measured = await root.evaluate(element => {
        const shellBounds = element.getBoundingClientRect();
        const cardElement = element.querySelector('[data-web-part-shell-card]');
        const cardBounds = cardElement.getBoundingClientRect();
        const header = element.querySelector('[data-web-part-shell-header]');
        const inside = header.parentElement === cardElement;
        const css = getComputedStyle(cardElement);
        return {
          left: cardBounds.left - shellBounds.left,
          right: shellBounds.right - cardBounds.right,
          bottom: shellBounds.bottom - cardBounds.bottom,
          top: (inside ? cardBounds : header.getBoundingClientRect()).top - shellBounds.top,
          inner: [css.paddingTop, css.paddingRight, css.paddingBottom, css.paddingLeft],
          outsideHeaderGap: inside ? undefined : cardBounds.top - header.getBoundingClientRect().bottom,
        };
      });
      const name = `Card ${position} application ${width}px/${padding}`;
      await check(`${name}: border detached by outer surface padding`, [measured.left, measured.right, measured.top, measured.bottom], [expectedInset, expectedInset, expectedInset, expectedInset]);
      await check(`${name}: native inner padding stays distinct`, measured.inner, ['12px', '12px', '12px', '12px']);
      if (position === 'outside') await check(`${name}: top follows outside header and gap`, measured.outsideHeaderGap, 12);
      await retainedCardState(name);
    }
  }
  await render({ surface: undefined, padding: undefined, width: 900 });

  // Explicit fixture caller props override live demo choices, then release back
  // to the unchanged options held above the Shell's structural boundary.
  await render({ appearance: 'filled', size: 'small', titlePosition: 'outside' });
  await appearanceSelect().selectOption('filled-alternative');
  await positionSelect().selectOption('inside');
  await sizeSelect().selectOption('large');
  await check('Explicit consumer size wins over sample option', await card().evaluate(element => getComputedStyle(element).paddingTop), '8px');
  await check('Explicit consumer title placement wins over sample option', await cardHeader().evaluate(header => header.parentElement.hasAttribute('data-web-part-shell-content')), true);
  const forcedBackground = await card().evaluate(element => getComputedStyle(element).backgroundColor);
  await render({ appearance: undefined, size: undefined, titlePosition: undefined });
  await check('Releasing consumer size restores selected large', await card().evaluate(element => getComputedStyle(element).paddingTop), '16px');
  await sizeSelect().selectOption('medium');
  await check('Releasing consumer appearance restores selected alternative', await card().evaluate(element => getComputedStyle(element).backgroundColor) !== forcedBackground, true);
  await check('Releasing consumer placement restores selected inside', await cardHeader().evaluate(header => header.parentElement.hasAttribute('data-web-part-shell-card')), true);
  await render({ mode: 'Edit' });
  await editable.waitFor();
  await editable.evaluate(element => { window.__shellSizeEditor = element; });
  await root.evaluate(element => { window.__shellSizeEditorHeader = element.querySelector('[data-web-part-shell-header]'); });
  const sizeProperties = (await observations()).properties;
  for (const size of ['small', 'medium', 'large', 'medium']) {
    await sizeSelect().selectOption(size);
    await check(`Size ${size} retains editor DOM`, await editable.evaluate(element => element === window.__shellSizeEditor), true);
    await check(`Size ${size} retains editor header`, await cardHeader().evaluate(element => element === window.__shellSizeEditorHeader), true);
    await check(`Size ${size} does not write properties`, (await observations()).properties, sizeProperties);
    await retainedCardState(`Edit size ${size}`);
  }
  for (const dir of ['ltr', 'rtl']) {
    await page.setViewportSize({ width: 320, height: 900 });
    await render({ mode: 'Edit', width: 320, dir, properties: { title: 'Alpha Beta Gamma', titleHTML: '<h2>Alpha Beta Gamma</h2>' } });
    await editable.waitFor(); await selectRange(0, 5);
    await check(`Narrow Card ${dir}: title editor remains inside Card`, await editable.evaluate(element => Boolean(element.closest('[data-web-part-shell-card]'))), true);
    await check(`Narrow Card ${dir}: toolbar is portaled outside Card`, await toolbar().evaluate(element => !element.closest('[data-web-part-shell-card]')), true);
    const bounds = await geometry();
    await check(`Narrow Card ${dir}: toolbar fits viewport`, bounds.left >= -1 && bounds.right <= bounds.viewport + 1, true);
    await button('More text formatting').click();
    await check(`Narrow Card ${dir}: overflow menu opens`, await page.getByRole('menuitemcheckbox', { name: 'Underline', exact: true }).isVisible(), true);
    await page.getByRole('menuitemcheckbox', { name: 'Underline', exact: true }).click();
    await waitHTML('<u>'); await focusProof(`Narrow Card ${dir} underline`);
    await retainedCardState(`Narrow Card ${dir}`);
    await capture(`card-narrow-${dir}`);
    // Native checkbox commands retain the popup. Dismiss the complete editing
    // surface so the next direction starts with a closed More menu.
    await page.locator('[data-fixture-outside]').click();
    await toolbar().waitFor({ state: 'detached' });
  }
  await page.setViewportSize({ width: 1210, height: 900 });
  await render({ mode: 'Read', width: 900, dir: 'ltr' });
  await appearanceSelect().selectOption('subtle');
  await check('Returning sample to subtle removes native Card', await card().count(), 0);
  await check('Returning subtle keeps live appearance/placement choices', [await appearanceSelect().inputValue(), await positionSelect().inputValue()], ['subtle', 'inside']);
  await check('Structural Card/subtle transitions retain outer root', await root.evaluate(element => element === window.__shellCardOuterRoot), true);
  await check('Structural transitions retain independent instance', await other.getByRole('textbox', { name: 'Other child', exact: true }).inputValue(), 'Independent Card state');

  // Fresh page isolates the exact intentionally failed lazy request; normal-page
  // console/page errors remain clean. Retry requests the real chunk afterward.
  const retryContext = await page.context().browser().newContext({ viewport: { width: 1210, height: 900 } });
  const retryPage = await retryContext.newPage(); const injectedErrors = []; let aborted;
  retryPage.on('console', message => { if (['error', 'warning'].includes(message.type())) injectedErrors.push(message.text()); });
  retryPage.on('pageerror', error => injectedErrors.push(error.message));
  await retryPage.route('**/*editable-text-editor*js', async route => {
    if (!aborted) { aborted = route.request().url(); await route.abort('failed'); } else await route.continue();
  });
  try {
    await retryPage.goto(url); await retryPage.locator('[data-shell-root="first"] h2').waitFor();
    await retryPage.evaluate(() => window.shellFixture.render({ mode: 'Edit', otherProperties: { showTitle: false } }));
    await retryPage.getByRole('button', { name: 'Retry', exact: true }).waitFor();
    await retryPage.locator('[data-shell-demo="note"]').fill('Retry retained note');
    await retryPage.getByRole('button', { name: 'Retry', exact: true }).click();
    await retryPage.locator('[data-shell-root="first"] [contenteditable]').waitFor();
    await check('Loader retry mounts actual editor', await retryPage.locator('[contenteditable]').count(), 1);
    await check('Loader failure never remounts children', await retryPage.evaluate(() => window.shellFixture.observe().mounts), 1);
    await check('Loader retry retains child note', await retryPage.locator('[data-shell-demo="note"]').inputValue(), 'Retry retained note');
    await check('Injected loader failure was exact editor chunk', Boolean(aborted && aborted.includes('editable-text-editor')), true);
    await retryPage.screenshot({ path: path.join(output, `${label}-shell-loader-retry.png`) });
  } finally { await retryContext.close(); }
  await check('Retry scenario has only exact injected chunk diagnostics', injectedErrors.length > 0 && injectedErrors.every(message => message === 'Failed to load resource: net::ERR_FAILED' || (message.startsWith('Loading chunk editable-text-editor failed.') && message.includes(aborted)) || (message.startsWith('The above error occurred in one of your React components:') && message.includes('EditableTextEditorBoundary'))), true);
  fs.writeFileSync(path.join(output, `${label}-shell-input-boundaries.json`), JSON.stringify({
    browserRouting: 'Browser plugin not available', input: ['Chromium mouse clicks, keyboard typing/shortcuts/Shift range selection', 'DOM Selection API range selection with actual selectionchange', 'Synthetic ClipboardEvent using native DataTransfer', 'Chromium CDP Input.imeSetComposition + Input.insertText; not physical OS IME'],
    physicalIME: 'NOT EXECUTED', touchHardware: 'NOT EXECUTED', screenReader: 'NOT EXECUTED', nativeTenantCommandParity: 'NOT EXECUTED', tenantSaveReload: 'NOT EXECUTED', injectedFailure: { aborted, errors: injectedErrors }, requests, final: await observations()
  }, null, 2));
}
module.exports = { checkWebPartShellBrowser };
