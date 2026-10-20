const assert = require('node:assert/strict'), fs = require('node:fs'), http = require('node:http'), path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const root = path.resolve(__dirname, '..'), directory = path.join(root, 'temp/pickers-browser'), evidence = path.resolve(process.env.PICKERS_BROWSER_EVIDENCE || path.join(root, '.docs/maintenance/evidence/pickers'));
async function main() {
  fs.mkdirSync(evidence, { recursive: true });
  const server = http.createServer((request, response) => {
    const pathname = new URL(request.url, 'http://127.0.0.1').pathname;
    if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
    const file = { '/': 'index.html', '/fixture.js': 'fixture.js', '/fixture.js.map': 'fixture.js.map' }[pathname];
    if (!file) { response.writeHead(404).end(); return; }
    response.setHeader('Content-Type', file.endsWith('.html') ? 'text/html' : 'application/javascript'); response.end(fs.readFileSync(path.join(directory, file)));
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  let browser, page; const errors = [], checks = [], layouts = [];
  const check = (name, actual, expected) => { assert.deepEqual(actual, expected, name); checks.push(name); };
  try {
    browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
    page = await browser.newPage({ viewport: { width: 1280, height: 1000 } });
    page.on('pageerror', error => errors.push(error.message)); page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.goto(`http://127.0.0.1:${server.address().port}`);
    const cases = page.locator('[data-picker-case]'); await cases.last().waitFor(); check('56 actual compiled layout cases', await cases.count(), 56);
    const rows = () => page.locator('.fui-Option:visible, .fui-TagPickerOption:visible');
    // Slow actual source adapters must announce pending work rather than a false empty menu.
    await page.evaluate(() => { window.pickerFixture.state.readDelayMs = 120; });
    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      const subject = cases.filter({ has: page.locator(`h2:text-is("${name}-combobox-single-wide")`) });
      const slowInput = subject.locator('.fui-Combobox input');
      await slowInput.click();
      await page.getByRole('status').filter({ hasText: 'Caricamento…' }).waitFor();
      check(`${name} slow initial read reports loading instead of empty`, (await page.getByRole('status').allTextContents()).some(text => text.includes('Digita almeno')), false);
      await rows().first().waitFor(); await page.getByText('Caricamento…', { exact: true }).waitFor({ state: 'hidden' });
      await page.waitForFunction(() => window.pickerFixture.stats.pending === 0);
      const beforeRows = await rows().allTextContents();
      const beforeCalls = await page.evaluate(() => [window.pickerFixture.stats.sp, window.pickerFixture.stats.graph]);
      await slowInput.evaluate(node => { window.pickerRetainedInput = node; });
      await page.keyboard.press('Escape');
      await subject.locator('[data-case-action="cosmetic"]').click(); await subject.locator('[data-case-action="equivalent"]').click();
      await page.evaluate(() => {
        window.pickerStatusFrames = [];
        window.pickerStatusObserver = new MutationObserver(() => window.pickerStatusFrames.push([...document.querySelectorAll('[role="status"]')].filter(n => n.getBoundingClientRect().width).map(n => n.textContent).join(' ')));
        window.pickerStatusObserver.observe(document.body, { subtree: true, childList: true, characterData: true });
      });
      await slowInput.click(); await rows().first().waitFor();
      await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
      const frames = await page.evaluate(() => { window.pickerStatusObserver.disconnect(); return window.pickerStatusFrames; });
      check(`${name} cosmetic/equivalent/reopen retains input DOM`, await slowInput.evaluate(node => node === window.pickerRetainedInput), true);
      check(`${name} settled reopen retains menu rows`, await rows().allTextContents(), beforeRows);
      check(`${name} cosmetic/equivalent/reopen reuses transport reads`, await page.evaluate(() => [window.pickerFixture.stats.sp, window.pickerFixture.stats.graph]), beforeCalls);
      check(`${name} settled reopen has no loading/empty flash`, frames.some(text => /Caricamento|Digita almeno|Nessun risultato/.test(text)), false);
      await page.keyboard.press('Escape');
    }
    await page.evaluate(() => { window.pickerFixture.state.readDelayMs = 0; });
    for (let i = 0; i < 56; i++) {
      const card = cases.nth(i), id = await card.getAttribute('data-picker-case');
      const input = card.locator('.fui-Combobox input, .fui-TagPickerInput');
      await input.click(); await rows().first().waitFor();
      const expected = id.startsWith('ListPicker-') ? 30 : /^(Field|View|ContentType)Picker-/.test(id) ? 35 : 5;
      await page.waitForFunction(count => Array.from(document.querySelectorAll('.fui-Option, .fui-TagPickerOption')).filter(node => node.getBoundingClientRect().width).length === count, expected);
      check(`${id} bounded pre-search candidates`, await rows().count(), expected);
      const popup = page.locator('.fui-Listbox:visible, .fui-TagPickerList:visible');
      check(`${id} bounded popup height`, await popup.evaluate(node => node.getBoundingClientRect().height <= 322), true);
      const geometry = await rows().first().evaluate(row => {
        const persona = row.querySelector('.fui-Persona'), primary = row.querySelector('.fui-Persona__primaryText'), secondary = row.querySelector('.fui-Persona__secondaryText');
        const r = row.getBoundingClientRect(), p = persona.getBoundingClientRect(), text = primary.getBoundingClientRect(), second = secondary?.getBoundingClientRect();
        return { offset: Math.round(p.left - r.left - parseFloat(getComputedStyle(row).paddingLeft)), aligned: second === undefined || Math.abs(second.left - text.left) < 1, overflows: row.scrollWidth > row.clientWidth + 1 };
      });
      check(`${id} primary and secondary alignment`, geometry.aligned, true); check(`${id} no row overflow`, geometry.overflows, false);
      if (id.includes('combobox-single')) check(`${id} no invisible check column`, geometry.offset, 0);
      if (id.includes('combobox-multi')) check(`${id} native checkbox column`, geometry.offset >= 16, true);
      layouts.push({ id, geometry });
      if (i < 2 || id === 'PrincipalPicker-tagPicker-multi-narrow') await page.screenshot({ path: path.join(evidence, `${id}.png`) });
      await page.keyboard.press('Escape');
    }
    const card = id => page.locator(`[data-picker-case="${id}"]`), action = (subject, name) => subject.locator(`[data-case-action="${name}"]`);
    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      for (const presentation of ['combobox', 'tagPicker']) for (const mode of ['single', 'multi']) {
        const subject = card(`${name}-${presentation}-${mode}-wide`), queryInput = subject.locator('.fui-Combobox input, .fui-TagPickerInput');
        await queryInput.click(); await queryInput.fill('zzzznomatch');
        await page.getByRole('status').filter({ hasText: 'Nessun risultato per questa ricerca' }).waitFor();
        check(`${name}-${presentation}-${mode} completed empty query feedback`, await page.getByRole('status').last().textContent(), 'Nessun risultato per questa ricerca');
        await queryInput.fill('x');
        check(`${name}-${presentation}-${mode} below minimum prompt`, await page.getByRole('status').last().textContent(), 'Digita almeno 2 caratteri per cercare');
        await queryInput.fill(''); await page.keyboard.press('Escape');
      }
    }

    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      const subject = card(`${name}-combobox-single-wide`), input = subject.locator('.fui-Combobox input');
      await input.click(); await rows().first().click();
      await page.waitForFunction(id => document.querySelector(`[data-picker-case="${id}"] [data-case-selection]`).textContent !== '[]', `${name}-combobox-single-wide`);
      const selected = await subject.locator('[data-case-selection]').textContent();
      await input.click(); await page.getByText('Selezionato', { exact: true }).waitFor();
      const checked = page.locator('.fui-Option:visible[aria-selected="true"]');
      check(`${name} current single row checked once`, await checked.count(), 1);
      check(`${name} current single check visible`, await checked.locator('.fui-Option__checkIcon').count(), 1);
      await input.fill('unrelated'); await page.getByText('Selezionato', { exact: true }).waitFor();
      check(`${name} single selection retained during unrelated search`, await checked.count(), 1);
      await checked.click();
      check(`${name} clicking current value preserves selection`, await subject.locator('[data-case-selection]').textContent(), selected);
      await subject.locator('.fui-Combobox__clearIcon').click();
      await page.waitForFunction(id => document.querySelector(`[data-picker-case="${id}"] [data-case-selection]`).textContent === '[]', `${name}-combobox-single-wide`);
      check(`${name} single clearable retained`, JSON.parse(await subject.locator('[data-case-selection]').textContent()), []);
      await page.keyboard.press('Escape');
    }
    for (const name of ['FieldPicker', 'ViewPicker', 'ContentTypePicker']) {
      const subject = card(`${name}-combobox-single-wide`), input = subject.locator('.fui-Combobox input');
      await input.click(); await rows().last().waitFor();
      check(`${name} complete catalog beyond thirty`, await rows().count(), 35);
      const before = await page.evaluate(() => window.pickerFixture.stats.sp);
      const title = name === 'FieldPicker' ? 'Field 34' : name === 'ContentTypePicker' ? 'Content type 34' : 'View 34';
      await input.fill(title.toUpperCase());
      await page.waitForFunction(() => [...document.querySelectorAll('.fui-Option')].filter(n => n.getBoundingClientRect().width).length === 1);
      check(`${name} local case insensitive match`, (await rows().first().textContent()).includes(title), true);
      await input.fill(title.toLowerCase());
      check(`${name} typing does not call REST`, await page.evaluate(() => window.pickerFixture.stats.sp), before);
      await page.keyboard.press('Escape');
    }
    const list = card('ListPicker-combobox-multi-wide'), listInput = list.locator('.fui-Combobox input');
    await action(list, 'empty').click(); await listInput.click(); await rows().first().waitFor();
    check('metadata catalog default thirty', await rows().count(), 30);
    check('catalog lookahead hint', (await page.getByRole('status').last().textContent()).includes('Sono disponibili altri elementi'), true);
    await listInput.fill('LIST 34'); await rows().first().waitFor();
    await page.waitForFunction(() => [...document.querySelectorAll('.fui-Option')].filter(n => n.getBoundingClientRect().width).length === 1);
    check('case insensitive search beyond initial thirty', (await rows().first().textContent()).includes('List 34'), true);
    await rows().first().click(); await page.waitForFunction(() => document.querySelector('[data-picker-case="ListPicker-combobox-multi-wide"] [data-case-selection]').textContent !== '[]');
    await listInput.fill('unrelated'); await page.getByText('Selezionati', { exact: true }).waitFor();
    check('selected remains checked with unrelated query', await rows().first().getAttribute('aria-checked'), 'true');
    await rows().first().click(); await page.waitForFunction(() => document.querySelector('[data-picker-case="ListPicker-combobox-multi-wide"] [data-case-selection]').textContent === '[]');
    check('multi Combobox deselection without matching search', JSON.parse(await list.locator('[data-case-selection]').textContent()), []);
    await page.keyboard.press('Escape');
    const contextual = card('SitePicker-combobox-multi-wide'), contextInput = contextual.locator('.fui-Combobox input');
    await action(contextual, 'empty').click(); await action(contextual, 'context-sites').click(); await contextInput.click();
    await page.getByText('Sito corrente', { exact: true }).waitFor(); await page.getByText('Hub associato', { exact: true }).waitFor();
    check('Site contextual suggestions available without query', await rows().count(), 2);
    await contextInput.fill('NO MATCH');
    await page.getByText('Sito corrente', { exact: true }).waitFor();
    await page.getByText('Hub associato', { exact: true }).waitFor();
    check('Site fixed choices survive unrelated typed search', await rows().count(), 2);
    await page.getByRole('status').filter({ hasText: 'Nessun risultato per questa ricerca' }).waitFor();
    check('fixed choices do not mask empty query feedback', await page.getByRole('status').last().textContent(), 'Nessun risultato per questa ricerca');
    await contextInput.fill('Site 0');
    await page.getByRole('status').filter({ hasText: 'Nessun altro risultato per questa ricerca' }).waitFor();
    check('matching fixed site is not reported as zero matches', await page.getByRole('status').last().textContent(), 'Nessun altro risultato per questa ricerca');
    await contextInput.fill('');
    await page.keyboard.press('Escape'); await action(contextual, 'update').click(); await contextInput.click();
    check('full consumer set cannot hide contextual choices', await page.getByText('Sito corrente', { exact: true }).count(), 1);
    check('hub remains with full consumer set', await page.getByText('Hub associato', { exact: true }).count(), 1);
    check('contextual choices retain consumer model precedence', (await rows().filter({ hasText: 'Renamed Site 0' }).count()), 1);

    await rows().filter({ hasText: 'Site 0' }).first().click();
    await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-multi-wide"] [data-case-selection]').textContent !== '[]');
    check('Site current suggestion selectable and deduplicated', await page.locator('.fui-Option:visible[aria-checked="true"]').count(), 1);
    const contextRequests = await page.evaluate(() => window.pickerFixture.stats.sp);
    await page.keyboard.press('Escape'); await contextInput.click(); await page.getByText('Hub associato', { exact: true }).waitFor();
    check('Site suggestions reopen reuses requests', await page.evaluate(() => window.pickerFixture.stats.sp), contextRequests);
    await page.keyboard.press('Escape');
    const site = card('SitePicker-combobox-single-wide'), input = site.locator('.fui-Combobox input');
    await input.click(); await rows().first().click(); await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-single-wide"] [data-case-selection]').textContent !== '[]');
    const selection = await site.locator('[data-case-selection]').textContent(); check('site JSON reference selected', JSON.parse(selection)[0].kind, 'site');
    check('no separate clear-all button', await site.getByRole('button', { name: 'Svuota selezione', exact: true }).count(), 0);
    await site.locator('.fui-Combobox__clearIcon').click();
    await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-single-wide"] [data-case-selection]').textContent === '[]');
    check('native Combobox clear emits empty references', JSON.parse(await site.locator('[data-case-selection]').textContent()), []);
    await input.click(); await rows().first().click();
    await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-single-wide"] [data-case-selection]').textContent !== '[]');
    await input.evaluate(node => { window.pickerRetainedInput = node; });
    const before = await page.evaluate(() => window.pickerFixture.stats.sp);
    await action(site, 'update').click(); await input.click(); await rows().first().waitFor();
    check('consumer update retains input DOM', await input.evaluate(node => node === window.pickerRetainedInput), true);
    check('consumer update retains selection', await site.locator('[data-case-selection]').textContent(), selection);
    check('renamed consumer option updates menu', (await rows().first().textContent()).includes('Renamed'), true);
    check('cosmetic consumer models reuse eligibility requests', await page.evaluate(() => window.pickerFixture.stats.sp), before);
    await page.keyboard.press('Escape'); await action(site, 'empty').click(); await action(site, 'clear').click(); await input.click();
    check('empty set prompts for two characters', await page.getByRole('status').last().textContent(), 'Digita almeno 2 caratteri per cercare');
    await input.fill('Site'); await rows().first().waitFor(); check('typed search includes nonconsumer sources', await rows().count(), 7);
    await page.keyboard.press('ArrowDown'); await page.keyboard.press('Enter');
    await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-single-wide"] [data-case-selection]').textContent !== '[]');
    check('keyboard selection confirms domain value', JSON.parse(await site.locator('[data-case-selection]').textContent())[0].kind, 'site');
    await page.keyboard.press('Escape');
    const principal = card('PrincipalPicker-tagPicker-multi-wide'), principalInput = principal.locator('.fui-TagPickerInput');
    await action(principal, 'empty').click(); await action(principal, 'personal').click(); await principalInput.click(); await rows().first().waitFor();
    await page.getByText('Persone suggerite', { exact: true }).waitFor(); await page.getByText('I tuoi gruppi', { exact: true }).waitFor();
    check('personal sources overall cap', await rows().count() <= 5, true); check('directory roles excluded', (await rows().allTextContents()).some(text => text.includes('directoryRole')), false);
    await rows().filter({ hasText: 'User 1' }).first().click();
    await page.waitForFunction(() => document.querySelector('[data-picker-case="PrincipalPicker-tagPicker-multi-wide"] [data-case-selection]').textContent !== '[]');
    const userValue = JSON.parse(await principal.locator('[data-case-selection]').textContent())[0]; check('person aggregate id not promoted', userValue.entra.id.startsWith('person'), false);
    await action(principal, 'both').click(); await principal.getByRole('button', { name: 'Risolvi ID', exact: true }).click();
    await page.waitForFunction(() => JSON.parse(document.querySelector('[data-picker-case="PrincipalPicker-tagPicker-multi-wide"] [data-case-selection]').textContent)[0]?.sharePoint?.id === 2);
    check('explicit both ID resolution retains Entra and SP IDs', JSON.parse(await principal.locator('[data-case-selection]').textContent())[0].entra.id, userValue.entra.id);
    check('read-only resolution did not ensureUser', await page.evaluate(() => window.pickerFixture.stats.ensure), 0);
    await page.locator('[data-global-action="personal-fail"]').click(); await action(principal, 'refresh').click(); await action(principal, 'update').click(); await principalInput.click(); await rows().filter({ hasText: 'Renamed' }).first().waitFor();
    check('personal failure retains healthy consumer candidates', (await rows().allTextContents()).some(text => text.includes('Renamed')), true);
    await page.waitForFunction(() => document.querySelector('[data-picker-case="PrincipalPicker-tagPicker-multi-wide"] [data-case-errors]').textContent.includes('personal source failure'));
    check('consumer section precedes personal sections after settlement', (await rows().first().textContent()).includes('Renamed'), true);
    await page.keyboard.press('Escape');
    for (let index = 0; index < 48; index++) {
      const subject = cases.nth(index), id = await subject.getAttribute('data-picker-case');
      await action(subject, 'long').click();
      const longInput = subject.locator('.fui-Combobox input, .fui-TagPickerInput');
      await longInput.click(); await longInput.fill('');
      const longRow = rows().filter({ hasText: 'Nome molto lungo' }).first(); await longRow.waitFor();
      const bounds = await longRow.evaluate(row => {
        const text = row.querySelector('.fui-Persona__primaryText');
        const secondary = row.querySelector('.fui-Persona__secondaryText');
        return { truncated: text.scrollWidth > text.clientWidth, ellipsis: getComputedStyle(text).textOverflow, fullText: text.title === text.textContent, secondaryTruncated: !!secondary && secondary.scrollWidth > secondary.clientWidth && getComputedStyle(secondary).textOverflow === 'ellipsis',
          contained: text.getBoundingClientRect().right <= row.getBoundingClientRect().right - 4,
          secondaryContained: !secondary || secondary.getBoundingClientRect().right <= row.getBoundingClientRect().right - 4 };
      });
      check(`${id} long label actually truncated`, bounds.truncated, true);
      check(`${id} long label uses ellipsis`, bounds.ellipsis, 'ellipsis');
      check(`${id} full label remains available on hover`, bounds.fullText, true);
      if (id.startsWith('PrincipalPicker')) check(`${id} long metadata actually truncates`, bounds.secondaryTruncated, true);
      check(`${id} primary text stays inside row`, bounds.contained, true);
      check(`${id} secondary text stays inside row`, bounds.secondaryContained, true);
      if (id.includes('tagPicker')) {
        await longRow.click(); await page.keyboard.press('Escape');
        const tag = subject.locator('.fui-Tag').filter({ hasText: 'Nome molto lungo' }).first(); await tag.waitFor();
        const tagBounds = await tag.evaluate(node => {
          const text = node.querySelector('.fui-Tag__primaryText'), control = node.closest('.fui-TagPickerControl'), dismiss = node.querySelector('.fui-Tag__dismissIcon');
          return { contained: node.getBoundingClientRect().width <= control.getBoundingClientRect().width - 2,
            truncated: text.scrollWidth > text.clientWidth && getComputedStyle(text).textOverflow === 'ellipsis', dismissVisible: !!dismiss && dismiss.getBoundingClientRect().width > 0 };
        });
        check(`${id} long selected tag stays inside control`, tagBounds.contained, true);
        check(`${id} selected tag text uses ellipsis`, tagBounds.truncated, true);
        check(`${id} selected tag removal stays visible`, tagBounds.dismissVisible, true);
      }
      await page.keyboard.press('Escape');
    }
    // Real native factories mount into a separate SPFx-owned root only after gallery acceptance.
    console.log(`Gallery acceptance before native fixture: ${checks.length} checks, ${layouts.length} layouts.`);
    const galleryCheckCount = checks.length + 1, nativeChecks = [], nativeReads = [], nativeGeometry = [];
    const nativeCheck = (name, actual, expected) => { assert.deepEqual(actual, expected, name); nativeChecks.push(name); };
    const native = page.locator('[data-native-picker-fixture]');
    const skeleton = native.locator('[data-property-pane-picker-control]');
    const overlay = native.locator('[data-auto-skeleton-overlay]');
    const nativeInput = native.locator('.fui-TagPickerInput');
    const frames = async () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    const mountNative = async (name, mode = 'unresolved') => {
      await page.evaluate(({ name, mode }) => { window.pickerFixture.native.blockReads(); window.pickerFixture.native.mount(name, mode); }, { name, mode });
      await nativeInput.waitFor({ state: 'attached' }); await frames();
    };
    const releaseNative = async () => {
      await page.evaluate(() => window.pickerFixture.native.releaseReads());
      await page.waitForFunction(() => window.pickerFixture.stats.pending === 0);
      await page.waitForFunction(() => document.querySelector('[data-native-picker-fixture] [data-property-pane-picker-control]')?.getAttribute('aria-busy') === 'false');
      await frames();
    };
    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      const beforeReads = await page.evaluate(() => window.pickerFixture.native.requests.length);
      await mountNative(name);
      await overlay.waitFor();
      nativeCheck(`${name} initial unresolved hydration is busy`, await skeleton.getAttribute('aria-busy'), 'true');
      nativeCheck(`${name} label stays visible during skeleton`, await native.getByText(`Native ${name}`, { exact: true }).isVisible(), true);
      nativeCheck(`${name} hint stays visible during skeleton`, await native.getByText('Native hydration hint', { exact: true }).isVisible(), true);
      nativeCheck(`${name} input geometry is forced rectangle`, await native.locator('[data-skeleton-shape="rectangle"]').count(), 1);
      nativeCheck(`${name} actual overlay contains rectangle geometry`, await overlay.locator('[data-auto-skeleton-shape="rectangle"]').count() > 0, true);
      const nativeBounds = await native.evaluate(node => {
        const label = node.querySelector('.fui-Field__label').getBoundingClientRect();
        const hint = node.querySelector('.fui-Field__hint').getBoundingClientRect();
        const control = node.querySelector('.fui-TagPickerControl').getBoundingClientRect();
        const overlay = node.querySelector('[data-auto-skeleton-overlay]').getBoundingClientRect();
        const shape = node.querySelector('[data-auto-skeleton-shape="rectangle"]').getBoundingClientRect();
        return { control: { width: control.width, height: control.height }, shape: { width: shape.width, height: shape.height },
          coversControl: shape.width >= control.width - 2 && shape.height >= control.height - 2,
          excludesLabelAndHint: overlay.top >= label.bottom - 1 && overlay.bottom <= hint.top + 1 };
      });
      nativeGeometry.push({ name, ...nativeBounds });
      nativeCheck(`${name} skeleton covers the input geometry`, nativeBounds.coversControl, true);
      nativeCheck(`${name} skeleton excludes label and hint geometry`, nativeBounds.excludesLabelAndHint, true);
      nativeCheck(`${name} unresolved reference has no text before transport`, (await native.locator('.fui-Tag__primaryText').allTextContents()).some(text => /Site 0|List 0|Field 0|View 0|Content type 0|Contract.pdf|User 0/.test(text)), false);
      await nativeInput.evaluate(node => { window.nativeRetainedInput = node; });
      const references = await page.evaluate(() => window.pickerFixture.native.values.current);
      if (name === 'SitePicker') await native.screenshot({ path: path.join(evidence, 'native-initial-hydration.png') });
      await releaseNative();
      nativeCheck(`${name} input DOM survives hydration`, await nativeInput.evaluate(node => node === window.nativeRetainedInput), true);
      nativeCheck(`${name} overlay ends after hydration`, await overlay.count(), 0);
      nativeCheck(`${name} hydrated selection becomes visible`, (await native.locator('.fui-Tag__primaryText').allTextContents()).some(text => /Site 0|List 0|Field 0|View 0|Content type 0|Contract.pdf|User 0/.test(text)), true);
      nativeCheck(`${name} initial hydration does not emit changes`, await page.evaluate(() => window.pickerFixture.native.changes.length), 0);
      nativeCheck(`${name} reference stays unchanged`, await page.evaluate(() => window.pickerFixture.native.values.current), references);
      const hydrationRequests = await page.evaluate(start => window.pickerFixture.native.requests.slice(start), beforeReads);
      nativeReads.push({ name, requests: hydrationRequests, count: hydrationRequests.length });
      nativeCheck(`${name} hydration has bounded real transport requests`, hydrationRequests.length > 0 && hydrationRequests.length <= 6, true);
      await page.evaluate(() => {
        window.nativeBusyFrames = [];
        window.nativeBusyObserver = new MutationObserver(() => window.nativeBusyFrames.push(document.querySelector('[data-native-picker-fixture] [data-property-pane-picker-control]')?.getAttribute('aria-busy')));
        window.nativeBusyObserver.observe(document.querySelector('[data-native-picker-fixture]'), { subtree: true, attributes: true, childList: true });
        window.pickerFixture.native.update('equivalent');
      });
      await frames();
      nativeCheck(`${name} equivalent props reuse hydration requests`, await page.evaluate(() => window.pickerFixture.native.requests.length), beforeReads + hydrationRequests.length);
      await nativeInput.click(); await nativeInput.fill('zzzznomatch');
      await page.getByRole('status').filter({ hasText: 'Nessun risultato per questa ricerca' }).last().waitFor();
      await page.keyboard.press('Escape');
      await page.evaluate(() => { window.pickerFixture.native.blockReads(); window.pickerFixture.native.update('refresh'); window.pickerFixture.native.update('selection'); });
      await frames();
      nativeCheck(`${name} changed reference keeps input DOM`, await nativeInput.evaluate(node => node === window.nativeRetainedInput), true);
      nativeCheck(`${name} subsequent hydration does not restart skeleton`, await skeleton.getAttribute('aria-busy'), 'false');
      await releaseNative();
      const busyFrames = await page.evaluate(() => { window.nativeBusyObserver.disconnect(); return window.nativeBusyFrames; });
      nativeCheck(`${name} search/refresh/equivalent/selection never reactivate skeleton`, busyFrames.includes('true'), false);
      await page.evaluate(() => window.pickerFixture.native.dispose());
      await mountNative(name); await overlay.waitFor();
      nativeCheck(`${name} true remount starts a fresh skeleton`, await skeleton.getAttribute('aria-busy'), 'true');
      await releaseNative();
      await page.evaluate(() => window.pickerFixture.native.dispose());
      for (const mode of ['label', 'model', 'availableModel', 'mixed', 'empty']) {
        await mountNative(name, mode);
        nativeCheck(`${name} ${mode} initial state bypasses skeleton`, await skeleton.getAttribute('aria-busy'), 'false');
        nativeCheck(`${name} ${mode} has no overlay`, await overlay.count(), 0);
        if (mode === 'label' || mode === 'mixed') nativeCheck(`${name} ${mode} retains readable snapshot`, await native.getByText('Retained display snapshot', { exact: true }).isVisible(), true);
        if (mode === 'mixed') nativeCheck(`${name} mixed unknown label has loading text`, await native.getByText('Caricamento…', { exact: true }).isVisible(), true);
        await releaseNative(); await page.evaluate(() => window.pickerFixture.native.dispose());
      }
    }
    for (const name of ['SitePicker', 'PrincipalPicker']) {
      await page.evaluate(name => { window.pickerFixture.native.blockReads(); window.pickerFixture.native.mount(name, 'unresolved', 'combobox'); }, name);
      const comboInput = native.locator('.fui-Combobox input');
      await comboInput.waitFor({ state: 'attached' }); await overlay.waitFor();
      nativeCheck(`${name} native Combobox uses rectangle geometry`, await overlay.locator('[data-auto-skeleton-shape="rectangle"]').count() > 0, true);
      nativeCheck(`${name} native Combobox label remains visible`, await native.getByText(`Native ${name}`, { exact: true }).isVisible(), true);
      nativeCheck(`${name} native Combobox hint remains visible`, await native.getByText('Native hydration hint', { exact: true }).isVisible(), true);
      await comboInput.evaluate(node => { window.nativeRetainedInput = node; });
      await releaseNative();
      nativeCheck(`${name} native Combobox keeps input DOM after hydration`, await comboInput.evaluate(node => node === window.nativeRetainedInput), true);
      nativeCheck(`${name} native Combobox displays hydrated value`, await comboInput.inputValue(), name === 'SitePicker' ? 'Site 0' : 'User 0');
      await page.evaluate(() => window.pickerFixture.native.dispose());
    }
    await mountNative('ListPicker'); await overlay.waitFor();
    await page.evaluate(() => { window.pickerFixture.state.fail = true; });
    await releaseNative();
    nativeCheck('terminal failure exits initial skeleton', await overlay.count(), 0);
    nativeCheck('terminal failure exposes native validation', await native.locator('.fui-Field__validationMessage').getByText('Simulated SharePoint failure', { exact: true }).isVisible(), true);
    nativeCheck('terminal failure exposes retry action', await native.getByRole('button', { name: 'Riprova', exact: true }).isVisible(), true);
    await nativeInput.click(); await page.getByRole('status').filter({ hasText: 'Simulated SharePoint failure' }).waitFor();
    nativeCheck('terminal hydration error is readable', await page.getByRole('status').filter({ hasText: 'Simulated SharePoint failure' }).isVisible(), true);
    await page.keyboard.press('Escape');
    await native.screenshot({ path: path.join(evidence, 'native-terminal-failure.png') });
    await page.evaluate(() => { window.pickerFixture.native.dispose(); window.pickerFixture.state.fail = false; });
    await page.setViewportSize({ width: 390, height: 1000 }); await page.screenshot({ path: path.join(evidence, 'narrow-viewport.png') });
    check('no browser errors', errors, []);
    fs.writeFileSync(path.join(evidence, 'verification.json'), JSON.stringify({ boundary: 'compiled', mock: 'SDK service identifiers and REST transports', checks, galleryCheckCount, nativeChecks, nativeReads, nativeGeometry, layouts, errors, stats: await page.evaluate(() => window.pickerFixture.stats) }, null, 2));
    console.log(`Compiled pickers browser: ${checks.length} gallery checks + ${nativeChecks.length} native checks, ${layouts.length} layouts, errors ${errors.length}. Evidence ${evidence}`);
  } catch (error) {
    if (page) { const diagnostic = await page.evaluate(() => ({ stats: window.pickerFixture.stats, state: window.pickerFixture.state, open: [...document.querySelectorAll('input[aria-expanded="true"]')].map(n => n.getAttribute('aria-label')), rows: [...document.querySelectorAll('[role="option"],[role="menuitemcheckbox"]')].filter(n => n.getBoundingClientRect().width).map(n => n.textContent), errors: document.querySelector('[data-picker-case="PrincipalPicker-tagPicker-multi-wide"] [data-case-errors]')?.textContent })); console.error(JSON.stringify(diagnostic)); await page.screenshot({ path: path.join(evidence, 'failure.png') }); }
    throw error;
  } finally { await browser?.close(); await new Promise(resolve => server.close(resolve)); }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
