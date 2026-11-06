// Actual explorer/source-module acceptance; transport ledgers remain separate from assets.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { scenario } = require('../../fixtures/executable-examples/cards-description-list.cjs');
const { parseContract } = require('../../../scripts/docs/documentation-contract.cjs');
const sourceDirectory = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples');
const italianCaption = /\b(?:Configurazione|Diagnostica|Titolo|Descrizione|Larghezza|Caricamento|Esegui|Salva|Ripristina|Operazione|Ordine|Ordini|Preferenze|Contatti|Aggiornamento|Immagine|Contenuto|Proprietà|Nessuna|Nessun)\b/u;

// Wait for real Fluent motion before accepting source geometry/screenshots.
async function settleSource(code) {
  await code.evaluate(async element => {
    const animations = element.closest('.fui-Accordion').getAnimations({ subtree: true });
    await Promise.all(animations.filter(animation => animation.effect.getTiming().iterations !== Infinity).map(animation => animation.finished.catch(() => undefined)));
  });
}

// Exclude only literal text; viewer chrome and accessible attributes stay checked.
async function interfaceCopy(panel) {
  return panel.evaluate(element => {
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
    const text = []; let node;
    while ((node = walker.nextNode())) if (!node.parentElement.closest('[data-example-source-code]')) text.push(node.textContent);
    for (const field of element.querySelectorAll('input,textarea,[aria-label],[title]')) {
      text.push(field.value || '', field.getAttribute('aria-label') || '', field.getAttribute('title') || '');
    }
    return text.join(' ');
  });
}
function authored(title) {
  const selected = scenario(title);
  const ast = ts.createSourceFile('source.ts', fs.readFileSync(path.join(sourceDirectory, `${selected.entry.id}.source.ts`), 'utf8'), ts.ScriptTarget.Latest, true);
  const declaration = ast.statements[0].declarationList.declarations[0];
  assert.equal(declaration.name.text, 'code'); assert.ok(ts.isStringLiteral(declaration.initializer));
  const code = declaration.initializer.text;
  const { file, line } = selected.entry.exampleSource;
  const physical = fs.readFileSync(path.resolve(__dirname, '../../..', file), 'utf8');
  const offset = physical.split('\n').slice(0, line - 1).join('\n').length + 1;
  const start = physical.lastIndexOf('/**', offset);
  assert.ok(start >= 0);
  const end = physical.indexOf('*/', start);
  assert.ok(end > offset);
  const diagnostics = [];
  const contract = parseContract(physical.slice(start, end + 2), { file, line: physical.slice(0, start).split('\n').length }, selected.entry.ownerKey, diagnostics);
  assert.deepEqual(diagnostics, []);
  assert.equal(contract.examples.find(example => example.source.line === line)?.code, code);
  return { ...selected, code };
}
const businessCounts = page => page.evaluate(() => ({
  clients: window.clientsFixture?.requests.length || 0,
  trim: window.trimFixture?.calls.length || 0,
  list: window.pnpReadFixture?.calls.length || 0,
  context: window.contextFixture?.requests.length || 0,
  properties: window.propertiesFixture?.requests.length || 0,
}));

async function exampleSourceBrowser({ page, check, select, activePanel, settle, overflow, capture, evidence, requests }) {
  evidence.assets = requests;
  const choose = async (key, title) => {
    await select(page, key);
    const panel = activePanel(page);
    const selector = panel.locator('select[aria-label="Scenario"]');
    if (await selector.count()) await selector.selectOption({ label: title });
    await panel.getByRole('heading', { name: title, exact: true }).waitFor();
    await panel.locator('[data-demo-section]').waitFor(); await settle(page);
    return panel;
  };
  const open = async (panel, expected, name) => {
    const header = panel.getByRole('button', { name: 'Example code', exact: true });
    check(`${name}: source starts closed`, await header.getAttribute('aria-expanded'), 'false');
    const before = requests.length;
    check(`${name}: matching source is not eagerly fetched`, requests.some(value => value.includes(expected.entry.id)), false);
    await header.click();
    const code = panel.locator('[data-example-source-code]'); await code.waitFor(); await settleSource(code);
    check(`${name}: literal source matches complete module.code`, await code.textContent(), expected.code);
    check(`${name}: matching source chunk fetched on disclosure`, requests.slice(before).map(value => path.basename(new URL(value).pathname)), [`spfx-jsdoc-source-${expected.entry.id}.js`]);
    check(`${name}: source provenance`, (await panel.textContent()).includes(`${expected.entry.exampleSource.file}:${expected.entry.exampleSource.line}`), true);
    check(`${name}: source stays outside actual demo`, await panel.locator('[data-demo-section] [data-example-source-code]').count(), 0);
    return { header, code };
  };
  const copy = async (panel, name, expected) => {
    await panel.getByRole('button', { name: 'Copy code', exact: true }).click();
    await panel.getByRole('status').filter({ hasText: /^(Copied|Unable to copy code\. Select and copy it manually\.)$/ }).waitFor();
    const status = await panel.getByRole('status').filter({ hasText: /^(Copied|Unable to copy code\.)/ }).textContent();
    evidence.clipboard.push({ name, outcome: status, boundary: 'Actual containing-document Clipboard API; no stub or permission grant' });
    if (status === 'Copied') {
      // Native paste is a clipboard readback without asking for read permission.
      const input = await page.evaluateHandle(() => { const field = document.createElement('textarea'); field.setAttribute('aria-label', 'Clipboard readback probe'); document.body.appendChild(field); field.focus(); return field; });
      try {
        await page.keyboard.press(process.platform === 'darwin' ? 'Meta+V' : 'Control+V');
        const value = await input.evaluate(field => field.value);
        if (value) { check(`${name}: actual OS clipboard paste readback is exact`, value, expected.code); evidence.clipboard.at(-1).readback = 'PASS: native paste'; }
        else evidence.clipboard.at(-1).readback = 'BLOCKED: ephemeral headless native paste supplied no text';
      } finally { await input.evaluate(field => field.remove()); await input.dispose(); }
    } else evidence.clipboard.at(-1).readback = 'BLOCKED: actual clipboard write unavailable/denied';
  };
  evidence.clipboard = [];
  check('Source: initial page fetched no source assets', requests.length, 0);
  const representatives = [
    ['cards', 'Inspect preview media and image recovery'],
    ['pnp-list-service', 'Read and page an explicitly selected list service'],
    ['web-part-shell', 'Configure the actual web part Shell'],
    ['native-site-source', 'Select native site references'],
  ];
  for (const [key, title] of representatives) {
    const expected = authored(title), panel = await choose(key, title);
    const draft = key === 'cards' ? panel.locator('[data-cards-demo=title-input]')
      : key === 'pnp-list-service' ? panel.getByRole('textbox', { name: 'List title', exact: true }) : undefined;
    if (draft) await draft.fill('Source action draft');
    const subject = await panel.locator('[data-demo-section]').elementHandle();
    const before = await businessCounts(page);
    const { header, code } = await open(panel, expected, key);
    await copy(panel, key, expected);
    await header.click(); await settle(page);
    check(`${key}: retained closed source is inert`, await code.evaluate(element => Boolean(element.closest('[inert]'))), true);
    await code.focus(); check(`${key}: inert source cannot receive focus`, await code.evaluate(element => document.activeElement === element), false);
    const count = requests.length; await header.click(); await code.waitFor();
    check(`${key}: successful source reopens without another request`, requests.length, count);
    check(`${key}: source actions retain actual subject DOM`, await subject.evaluate(element => element.isConnected), true);
    if (draft) check(`${key}: source actions retain current draft`, await draft.inputValue(), 'Source action draft');
    if (key === 'native-site-source') check('Native-adapter local boundary is an unconnected prerequisite', (await panel.textContent()).includes('The initialized native scenario host is not connected.'), true);
    check(`${key}: source actions make no business I/O`, await businessCounts(page), before);
    check(`${key}: English source chrome`, italianCaption.test(await interfaceCopy(panel)), false);
    await subject.dispose();
  }
  // A literal caption inside the source region is not part of the interface.
  // An actual caption just outside it must still fail the English scan.
  let panel = await choose('editable-text', 'Edit a controlled block with policy and dictionaries');
  const localized = authored('Edit a controlled block with policy and dictionaries');
  const { code: localizedCode } = await open(panel, localized, 'source localization');
  await localizedCode.evaluate(element => element.append('\n// Salva'));
  check('Source region contains the literal Italian caption', italianCaption.test(await localizedCode.textContent()), true);
  check('Source localization does not classify payload as interface', italianCaption.test(await interfaceCopy(panel)), false);
  await localizedCode.evaluate((element, text) => { element.textContent = text; }, localized.code);
  const injected = await panel.evaluateHandle(element => { const button = document.createElement('button'); button.textContent = 'Salva'; element.appendChild(button); return button; });
  check('English scan catches deliberate Italian UI outside source', italianCaption.test(await interfaceCopy(panel)), true);
  await injected.evaluate(element => element.remove()); await injected.dispose();
  await page.setViewportSize({ width: 320, height: 1000 });
  await page.evaluate(() => window.demoFixtureHost.update({ theme: 'dark', dir: 'rtl' })); await settle(page);
  const code = panel.locator('[data-example-source-code]'); await settleSource(code);
  await overflow(page, 'Example source dark RTL 320');
  check('Source code aligns authored indentation left in RTL host', await code.evaluate(element => {
    const css = getComputedStyle(element);
    return css.textAlign === 'start' ? css.direction === 'ltr' ? 'left' : 'right'
      : css.textAlign === 'end' ? css.direction === 'ltr' ? 'right' : 'left' : css.textAlign;
  }), 'left');
  check('Source code remains LTR in RTL host', await code.evaluate(element => getComputedStyle(element).direction), 'ltr');
  check('Source uses inherited dark surface token', await code.evaluate(element => getComputedStyle(element).getPropertyValue('--colorNeutralBackground2').trim()), require('@fluentui/react-components').teamsDarkTheme.colorNeutralBackground2);
  check('Source has a painted theme surface', await code.evaluate(element => getComputedStyle(element).backgroundColor !== 'rgba(0, 0, 0, 0)'), true);
  check('Source long lines scroll within bounded region', await code.evaluate(element => element.scrollWidth > element.clientWidth && element.clientHeight > 0 && element.clientHeight <= 384), true);
  await code.focus(); await page.keyboard.press('ArrowRight');
  await page.waitForFunction(() => document.querySelector('[data-example-source-code]').scrollLeft > 0);
  check('Source scroll region supports keyboard focus', await code.evaluate(element => document.activeElement === element), true);
  await capture(page, 'source-dark-rtl-320');
  await page.evaluate(() => window.demoFixtureHost.update({ theme: 'light', dir: 'ltr' }));
  await page.setViewportSize({ width: 1280, height: 1000 }); await settle(page);
  // Pending demo work is independent of disclosure/copy, then Reset preserves
  // immutable source state while restoring the existing demo baseline.
  panel = await choose('pagination', 'Browse controlled rows with pending completion');
  const pending = authored('Browse controlled rows with pending completion');
  await panel.getByRole('button', { name: 'Next page', exact: true }).click();
  const subject = await panel.locator('[data-demo-section]').elementHandle(), before = await businessCounts(page);
  await open(panel, pending, 'pending pagination'); await copy(panel, 'pending pagination', pending);
  check('Source leaves pending page locked', await panel.getByRole('button', { name: 'Next page', exact: true }).isDisabled(), true);
  check('Source leaves pending completion available', await panel.getByRole('button', { name: 'Complete pending request', exact: true }).isDisabled(), false);
  await panel.getByRole('button', { name: 'Reset scenario', exact: true }).click(); await settle(page);
  check('Reset retains immutable open source', await panel.locator('[data-example-source-code]').textContent(), pending.code);
  check('Reset retains demo mount', await subject.evaluate(element => element.isConnected), true); await subject.dispose();
  check('Pending/Reset source actions make no business I/O', await businessCounts(page), before);
  // Existing actual Shell retained updates versus subtle/Card parent remount.
  panel = await choose('web-part-shell', 'Retain content through Shell presentation changes');
  await panel.getByRole('button', { name: 'Use Card baseline', exact: true }).click(); await settle(page);
  const retained = authored('Retain content through Shell presentation changes');
  const note = panel.getByRole('textbox', { name: 'Project note', exact: true }); await note.fill('Source retained Shell draft');
  const identity = await note.elementHandle();
  const retainedOpen = await open(panel, retained, 'retained Shell');
  await copy(panel, 'retained Shell', retained);
  await retainedOpen.header.click(); await retainedOpen.header.click();
  check('Shell source actions preserve the edited draft', await identity.evaluate(element => element.isConnected && element.value === 'Source retained Shell draft'), true);
  await panel.getByRole('combobox', { name: 'Card appearance', exact: true }).selectOption('filled'); await settle(page);
  check('Retained Shell presentation preserves demo draft/node', await identity.evaluate(element => element.isConnected && element.value === 'Source retained Shell draft'), true);
  check('Retained Shell presentation preserves source disclosure', await panel.getByRole('button', { name: 'Example code', exact: true }).getAttribute('aria-expanded'), 'true'); await identity.dispose();
  panel = await choose('web-part-shell', 'Observe content remount between subtle and Card');
  await panel.getByRole('button', { name: 'Use subtle baseline', exact: true }).click(); await settle(page);
  const remount = authored('Observe content remount between subtle and Card'); await open(panel, remount, 'remounted Shell');
  const oldCode = await panel.locator('[data-example-source-code]').elementHandle();
  await panel.getByRole('button', { name: 'Use Card baseline', exact: true }).click(); await settle(page);
  check('Documented subtle/Card transition remounts source', await oldCode.evaluate(element => element.isConnected), false); await oldCode.dispose();
  check('Documented subtle/Card transition resets disclosure', await panel.getByRole('button', { name: 'Example code', exact: true }).getAttribute('aria-expanded'), 'false');
  await panel.getByRole('button', { name: 'Example code', exact: true }).click();
  check('Remounted Shell reloads same authored code', await panel.locator('[data-example-source-code]').textContent(), remount.code);
  // Deliberately deny only this previously unloaded source chunk. Chromium's
  // exact resource error is recorded separately; arbitrary diagnostics still fail.
  panel = await choose('relative-time', 'Explore relative clocks and fallback slots');
  const retry = authored('Explore relative clocks and fallback slots');
  const pattern = `**/spfx-jsdoc-source-${retry.entry.id}.js`;
  const retrySubject = await panel.locator('[data-demo-section]').elementHandle();
  const retryBefore = await businessCounts(page);
  let attempts = 0;
  await page.route(pattern, async route => {
    attempts++;
    if (attempts === 1) {
      evidence.deniedAsset = route.request().url();
      return route.abort('failed');
    }
    return route.continue();
  });
  const retryHeader = panel.getByRole('button', { name: 'Example code', exact: true });
  await retryHeader.click();
  await panel.getByRole('status').filter({ hasText: /^Unable to load example code\.$/ }).waitFor();
  check('Denied source: Copy remains disabled', await panel.getByRole('button', { name: 'Copy code', exact: true }).isDisabled(), true);
  await retryHeader.click(); await retryHeader.click(); await settle(page);
  check('Denied source: collapse/reopen does not retry', attempts, 1);
  await panel.getByRole('button', { name: 'Retry', exact: true }).click();
  await panel.locator('[data-example-source-code]').waitFor();
  check('Explicit Retry loads exact code when local bundler permits', await panel.locator('[data-example-source-code]').textContent(), retry.code);
  check('Explicit Retry reissues one source request', attempts, 2);
  check('Denied source/Retry preserve subject DOM', await retrySubject.evaluate(element => element.isConnected), true); await retrySubject.dispose();
  check('Denied source/Retry issue no business I/O', await businessCounts(page), retryBefore);
  await page.unroute(pattern);
  // This deliberate denial checks fallback behavior only; actual clipboard
  // evidence above continues to use the browser's unmodified implementation.
  await page.evaluate(() => Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined }));
  await panel.getByRole('button', { name: 'Copy code', exact: true }).click();
  await panel.getByRole('status').filter({ hasText: /^Unable to copy code\. Select and copy it manually\.$/ }).waitFor();
  check('Unavailable clipboard keeps complete manually selectable code', await panel.locator('[data-example-source-code]').textContent(), retry.code);
  check('Viewer failure chrome remains in English scan', italianCaption.test(await interfaceCopy(panel)), false);
  await page.evaluate(() => delete navigator.clipboard);
  evidence.simulatedClipboardFailure = 'PASS: unsupported API fallback only; separate from actual clipboard/readback';
  evidence.retry = { attempts, outcome: 'PASS: local webpack retry after one aborted source request; host restrictions not bypassed' };
  evidence.representatives = representatives.map(([key, title]) => ({ key, title, id: authored(title).entry.id }));
}
module.exports = { exampleSourceBrowser, interfaceCopy, italianCaption, authored, settleSource };
