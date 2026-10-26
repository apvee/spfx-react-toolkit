// Real browser acceptance shared by the existing source and compiled Shell runners.
// Commands use native keyboard, clipboard and Fluent UI, never engine handles.
const fs = require('node:fs');
const path = require('node:path');
async function checkEditableTextBrowser(page, record) {
  const { url, output, label, check } = record;
  const settle = () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  const query = scenario => `${url}/?scenario=${scenario}`;
  const subject = page.locator('[data-editable-text-demo=preview]');
  const editor = subject.locator('[contenteditable=true]');
  const toolbar = () => page.locator('[data-title-toolbar]');
  const demo = key => page.locator(`[data-editable-text-demo=${key}]`);
  const stored = () => demo('stored').textContent();
  const focus = async () => { await editor.focus(); await toolbar().waitFor(); await settle(); };
  const command = async name => {
    const direct = toolbar().getByRole('button', { name, exact: true });
    if (await direct.isVisible()) await direct.click();
    else { await toolbar().getByRole('button', { name: 'More text formatting', exact: true }).click(); await page.getByRole('menuitem', { name, exact: true }).or(page.getByRole('menuitemcheckbox', { name, exact: true })).click(); }
  };
  const requests = []; const onRequest = request => { if (/\.js(?:\?|$)/.test(request.url())) requests.push(request.url()); };
  page.on('request', onRequest);
  await page.goto(query('editable-text')); await subject.locator('h2').waitFor();
  await check(`${label}: pure panel one subject`, await subject.count(), 1);
  await check(`${label}: standalone creates no Shell`, await page.locator('[data-web-part-shell]').count(), 0);
  await check(`${label}: standalone Read requests no engine chunk`, requests.filter(value => !value.endsWith('/fixture.js')), []);
  await demo('mode').selectOption('edit'); await editor.waitFor(); await focus();
  await check(`${label}: Edit loads emitted shared chunk`, requests.some(value => value.includes('editable-text-editor')), true);
  // Real keyboard range selection survives a direct Fluent popup command.
  await page.keyboard.press('Meta+a'); await command('Theme color');
  await page.locator('[data-title-color=red]').click(); await settle();
  await check(`${label}: palette formats selected text`, (await stored()).includes('wps-color-red'), true);
  await check(`${label}: palette restores editor focus`, await editor.evaluate(element => element.ownerDocument.activeElement === element), true);
  await command('Clear formatting'); await settle();
  await check(`${label}: clear removes marks`, await stored(), '<h2>Project update</h2>');
  await page.keyboard.press('Meta+z'); await settle();
  await check(`${label}: clear is undoable`, (await stored()).includes('wps-color-red'), true);
  await demo('preset').selectOption('paragraph'); await focus();
  await check(`${label}: p-only projects p retaining text`, await editor.locator('p').textContent(), 'Project update');
  await check(`${label}: p-only has no block selector`, await toolbar().getByRole('combobox', { name: 'Heading', exact: true }).count(), 0);
  await page.keyboard.press('Meta+a'); await page.keyboard.type('## '); await page.keyboard.type('Paragraph'); await settle();
  await page.keyboard.press('Meta+Alt+2'); await settle();
  await check(`${label}: p-only keyboard cannot introduce heading`, await editor.locator('h1,h2,h3,h4').count(), 0);
  await page.keyboard.press('Enter'); await page.keyboard.type('Second'); await settle();
  await check(`${label}: Enter produces hard break`, await editor.locator('p br').count(), 1);
  // Native clipboard: copy literal multiline from a textarea, then browser paste.
  await page.evaluate(() => { const input = document.createElement('textarea'); input.id = 'editable-clipboard'; input.value = 'Clipboard\nsecond line'; document.body.appendChild(input); });
  await page.locator('#editable-clipboard').focus(); await page.keyboard.press('Meta+a'); await page.keyboard.press('Meta+c');
  await focus(); await page.keyboard.press('Meta+a'); await page.keyboard.press('Meta+v');
  await page.waitForFunction(() => document.querySelector('[data-editable-text-demo=preview] [contenteditable]').textContent.includes('Clipboard'));
  await check(`${label}: native clipboard retains p-only`, await editor.locator('p').count(), 1);
  await check(`${label}: native multiline paste has break`, await editor.locator('br').count(), 1);
  const beforePolicy = await stored(); const callbacks = await demo('status').textContent();
  await demo('preset').selectOption('headings'); await settle();
  await check(`${label}: policy change retains source without callback`, await stored(), beforePolicy);
  await check(`${label}: policy change callback count unchanged`, (await demo('status').textContent()).split('Callbacks:')[1], callbacks.split('Callbacks:')[1]);
  await focus(); await page.keyboard.press('Meta+a'); await page.keyboard.type('Headings'); await page.keyboard.press('Meta+Alt+0'); await settle();
  await check(`${label}: headings-only shortcut cannot create p`, await editor.locator('p').count(), 0);
  await toolbar().getByRole('combobox', { name: 'Heading', exact: true }).click();
  await check(`${label}: headings-only menu excludes p`, await page.getByRole('option', { name: 'Normal', exact: true }).count(), 0);
  await page.getByRole('option', { name: 'Heading 4', exact: true }).click(); await settle();
  await check(`${label}: heading popup applies exact block`, await editor.locator('h4').textContent(), 'Headings');
  await check(`${label}: heading popup restores focus`, await editor.evaluate(element => element.ownerDocument.activeElement === element), true);
  for (const theme of ['light', 'dark', 'contrast']) {
    await page.evaluate(theme => window.demoFixtureHost.update({ theme }), theme); await focus();
    await check(`${label}: ${theme} inherited Fluent tokens`, await editor.evaluate(element => Boolean(getComputedStyle(element).getPropertyValue('--colorNeutralForeground1').trim())), true);
  }
  await page.emulateMedia({ forcedColors: 'active' }); await focus(); await command('Theme color');
  await check(`${label}: forced colors keeps named choices`, await page.locator('[data-title-color]').evaluateAll(elements => elements.length > 1 && elements.every(element => element.getAttribute('aria-label'))), true);
  await page.keyboard.press('Escape'); await page.emulateMedia({ forcedColors: 'none' });
  await page.setViewportSize({ width: 320, height: 900 }); await page.evaluate(() => window.demoFixtureHost.update({ dir: 'rtl', theme: 'light' })); await focus(); await page.waitForTimeout(350);
  await toolbar().getByRole('button', { name: 'More text formatting', exact: true }).click();
  await check(`${label}: narrow overflow headings exclude paragraph`, await page.getByRole('menuitemradio', { name: 'Normal', exact: true }).count(), 0);
  await page.keyboard.press('Escape'); await focus(); await command('Theme color');
  await check(`${label}: RTL narrow popup fits`, await page.locator('[data-title-color-picker]').evaluate(element => { const rect = element.getBoundingClientRect(); return rect.left >= -1 && rect.right <= innerWidth + 1; }), true);
  await page.keyboard.press('Escape'); await page.screenshot({ path: path.join(output, `${label}-editable-rtl-narrow.png`), fullPage: true });
  await page.setViewportSize({ width: 1280, height: 1000 });
  await page.goto(query('editable-text-regressions'));
  const paragraph = page.locator('[data-editable-isolation=paragraph] [contenteditable=true]');
  const heading = page.locator('[data-editable-isolation=headings] [contenteditable=true]');
  await paragraph.waitFor(); await heading.waitFor();
  await check(`${label}: two standalone editors plus title`, await page.locator('[contenteditable=true]').count(), 3);
  await check(`${label}: native root ref reached`, JSON.parse(await page.locator('[data-editable-isolation=native]').textContent()).hasRef, true);
  await paragraph.click(); await settle();
  await check(`${label}: native root event forwarded`, JSON.parse(await page.locator('[data-editable-isolation=native]').textContent()).clicks > 0, true);
  await check(`${label}: consumer class and style forwarded`, await page.locator('[data-editable-isolation=paragraph]').evaluate(element => element.classList.contains('editable-consumer-class') && getComputedStyle(element).borderTopWidth === '3px'), true);
  await check(`${label}: first locale Italian`, await toolbar().getAttribute('aria-label'), 'Controlli di formattazione del testo');
  await heading.focus(); await settle();
  await check(`${label}: independent French locale`, await toolbar().getAttribute('aria-label'), 'Commandes de mise en forme du texte');
  await page.keyboard.press('Meta+a'); await page.keyboard.press('Backspace'); await page.keyboard.type('### '); await page.keyboard.type('Excluded'); await settle();
  await check(`${label}: H2/H4 excluded marker never becomes H4`, await heading.locator('h4').count(), 0);
  await check(`${label}: excluded input remains literal`, await heading.textContent(), '### Excluded');
  await check(`${label}: other editor isolated`, await paragraph.textContent(), 'Paragraph source');
  await page.locator('[data-editable-isolation=narrow]').press('Enter'); await paragraph.focus(); await settle(); await page.waitForTimeout(350);
  await toolbar().getByRole('button', { name: 'Altre opzioni di formattazione del testo', exact: true }).click();
  await check(`${label}: p-only overflow has no selector`, await page.getByRole('menuitem', { name: 'Intestazione', exact: true }).count(), 0);
  await page.keyboard.press('Escape'); await heading.focus(); await settle(); await page.waitForTimeout(350);
  await toolbar().getByRole('button', { name: 'Autres options de mise en forme du texte', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Titre', exact: true }).click();
  await check(`${label}: discontinuous overflow lists exact policy`, await page.getByRole('menuitemradio').allTextContents(), ['Titre 2', 'Titre 4']);
  await page.getByRole('menuitemradio', { name: 'Titre 4', exact: true }).click(); await settle();
  await check(`${label}: overflow selects allowed H4 and retains literal marker`, await heading.locator('h4').textContent(), '### Excluded');
  await check(`${label}: overflow restores heading editor focus`, await heading.evaluate(element => element.ownerDocument.activeElement === element), true);

  // ownerDocument/portals are exercised through a same-origin iframe with real fixture bytes.
  await page.goto(query('editable-text'));
  await page.evaluate(src => { const iframe = document.createElement('iframe'); iframe.id = 'editable-frame'; iframe.width = '900'; iframe.height = '700'; iframe.src = src; document.body.appendChild(iframe); }, query('editable-text'));
  const frame = page.frameLocator('#editable-frame'); await frame.locator('[data-editable-text-demo=mode]').selectOption('edit');
  const framed = frame.locator('[data-editable-text-demo=preview] [contenteditable=true]'); await framed.waitFor(); await framed.focus();
  await check(`${label}: iframe engine ownerDocument`, await framed.evaluate(element => element.ownerDocument !== window.parent.document && element.ownerDocument.activeElement === element), true);
  await frame.locator('[data-title-toolbar]').getByRole('button', { name: 'Bold', exact: true }).click();
  await check(`${label}: iframe toolbar restores focus`, await framed.evaluate(element => element.ownerDocument.activeElement === element), true);
  fs.writeFileSync(path.join(output, `${label}-editable-input-boundaries.json`), JSON.stringify({ requests, input: ['Chromium native keyboard/clipboard', 'Fluent direct/overflow popup interaction', 'same-origin iframe'], physicalIME: 'NOT EXECUTED', screenReader: 'NOT EXECUTED', authenticatedSharePoint: 'NOT EXECUTED' }, undefined, 2));
  page.off('request', onRequest);
}
module.exports = { checkEditableTextBrowser };
