const assert = require('node:assert/strict'), fs = require('node:fs'), http = require('node:http'), path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const root = path.resolve(__dirname, '..'), directory = path.join(root, 'temp/pickers-browser'), evidence = path.resolve(process.env.PICKERS_BROWSER_EVIDENCE || path.join(root, '.docs/maintenance/evidence/pickers/dropdown'));
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
    const query = () => page.locator('.fui-SearchBox input:visible, .fui-TagPickerInput:focus');
    const rows = () => page.locator('.fui-Option:visible, .fui-TagPickerOption:visible');
    // Slow actual source adapters must announce pending work rather than a false empty menu.
    await page.evaluate(() => { window.pickerFixture.state.readDelayMs = 120; });
    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      const subject = cases.filter({ has: page.locator(`h2:text-is("${name}-combobox-single-wide")`) });
      const slowTrigger = subject.locator('.fui-Dropdown__button');
      await slowTrigger.click();
      await page.getByRole('status').filter({ hasText: 'Caricamento…' }).waitFor();
      check(`${name} slow initial read reports loading instead of empty`, (await page.getByRole('status').allTextContents()).some(text => text.includes('Digita almeno')), false);
      await rows().first().waitFor(); await page.getByText('Caricamento…', { exact: true }).waitFor({ state: 'hidden' });
      await page.waitForFunction(() => window.pickerFixture.stats.pending === 0);
      const beforeRows = await rows().allTextContents();
      const beforeCalls = await page.evaluate(() => [window.pickerFixture.stats.sp, window.pickerFixture.stats.graph]);
      await slowTrigger.evaluate(node => { window.pickerRetainedInput = node; });
      await page.keyboard.press('Escape');
      await subject.locator('[data-case-action="cosmetic"]').click(); await subject.locator('[data-case-action="equivalent"]').click();
      await page.evaluate(() => {
        window.pickerStatusFrames = [];
        window.pickerStatusObserver = new MutationObserver(() => window.pickerStatusFrames.push([...document.querySelectorAll('[role="status"]')].filter(n => n.getBoundingClientRect().width).map(n => n.textContent).join(' ')));
        window.pickerStatusObserver.observe(document.body, { subtree: true, childList: true, characterData: true });
      });
      await slowTrigger.click(); await rows().first().waitFor();
      await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
      const frames = await page.evaluate(() => { window.pickerStatusObserver.disconnect(); return window.pickerStatusFrames; });
      check(`${name} cosmetic/equivalent/reopen retains trigger DOM`, await slowTrigger.evaluate(node => node === window.pickerRetainedInput), true);
      check(`${name} settled reopen retains menu rows`, await rows().allTextContents(), beforeRows);
      check(`${name} cosmetic/equivalent/reopen reuses transport reads`, await page.evaluate(() => [window.pickerFixture.stats.sp, window.pickerFixture.stats.graph]), beforeCalls);
      check(`${name} settled reopen has no loading/empty flash`, frames.some(text => /Caricamento|Digita almeno|Nessun risultato/.test(text)), false);
      await page.keyboard.press('Escape');
    }
    await page.evaluate(() => { window.pickerFixture.state.readDelayMs = 0; });
    for (let i = 0; i < 56; i++) {
      const card = cases.nth(i), id = await card.getAttribute('data-picker-case');
      const input = card.locator('.fui-Dropdown__button, .fui-TagPickerInput');
      await input.click(); await rows().first().waitFor();
      if (id.includes('combobox')) {
        const search = page.getByRole('searchbox', { name: `${id}: cerca`, exact: true });
        check(`${id} search has its own accessible name`, await search.count(), 1);
        check(`${id} search and trigger IDs are distinct`, await search.getAttribute('id') !== await input.getAttribute('id'), true);
      }
      const expected = id.startsWith('ListPicker-') ? 30 : /^(Field|View|ContentType)Picker-/.test(id) ? 35 : 5;
      await page.waitForFunction(count => Array.from(document.querySelectorAll('.fui-Option, .fui-TagPickerOption')).filter(node => node.getBoundingClientRect().width).length === count, expected);
      check(`${id} bounded pre-search candidates`, await rows().count(), expected);
      const popup = page.locator(id.includes('combobox') ? '.fui-Dropdown__listbox:visible' : '.fui-TagPickerList:visible');
      const popupHeight = await popup.evaluate(node => node.getBoundingClientRect().height);
      check(`${id} bounded whole popup height (${popupHeight}px)`, popupHeight <= 322, true);
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
    // Production Dropdown regressions replacing the disposable public-slot spike.
    const frames = async () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    const outside = page.locator('[data-global-action="outside"]');
    const settle = async message => { await page.waitForFunction(() => window.pickerFixture.confirmation.pending); await page.evaluate(message => window.pickerFixture.confirmation.settle(message), message); await frames(); };
    const selectedValues = async subject => JSON.parse(await subject.locator('[data-case-selection]').textContent());
    for (const mode of ['single', 'multi']) {
      const subject = card(`SitePicker-combobox-${mode}-narrow`), trigger = subject.locator('.fui-Dropdown__button');
      const search = page.locator('.fui-SearchBox input:visible');
      const events = () => subject.locator('[data-case-events]').textContent();
      const forwardTarget = () => subject.evaluate(node => {
        const clear = node.querySelector('.fui-Dropdown__clearButton');
        const target = clear && clear.tabIndex >= 0 && clear.getBoundingClientRect().width > 0 ? clear : node.querySelector('[data-case-after]');
        return target === document.activeElement;
      });
      const backwardTarget = () => subject.locator('[data-case-before]').evaluate(node => node === document.activeElement);
      const open = async () => { await trigger.click(); await search.waitFor(); await frames(); };
      await action(subject, 'defer').click(); await open();
      check(`${mode} search has an independent DOM identity`, await search.getAttribute('id') !== await trigger.getAttribute('id'), true);
      check(`${mode} label targets only the selection trigger`, await subject.locator('.fui-Field__label').getAttribute('for'), await trigger.getAttribute('id'));
      check(`${mode} production search receives focus`, await search.evaluate(node => node === document.activeElement), true);
      check(`${mode} search is outside semantic list`, await search.evaluate(node => !node.closest('[role="listbox"], [role="menu"]')), true);
      check(`${mode} popup is body portal`, await search.evaluate(node => !node.closest('.fui-Dropdown')), true);
      check(`${mode} native aria-controls owns the list`, await trigger.evaluate(node => ['listbox', 'menu'].includes(document.getElementById(node.getAttribute('aria-controls'))?.getAttribute('role'))), true);
      const retained = await search.elementHandle();
      await search.fill('Site '); await page.keyboard.type('1');
      check(`${mode} editable space retained`, await search.inputValue(), 'Site 1');
      check(`${mode} query rerender preserves search node`, await search.evaluate((node, retained) => node === retained, retained), true);
      await rows().first().waitFor(); await frames();
      // On macOS Chrome Home may scroll a long document instead of moving a search caret.
      // Preserve the browser default; the platform text-edit shortcut proves caret editing separately.
      await search.evaluate(node => { node.addEventListener('keydown', event => { if (event.key === 'Home' || event.key === 'End') queueMicrotask(() => { window.pickerEditPrevented = event.defaultPrevented; }); }); });
      for (const key of ['Home', 'End']) {
        const active = await search.getAttribute('aria-activedescendant'); await search.press(key);
        check(`${mode} ${key} retains native editor behavior`, await page.evaluate(() => window.pickerEditPrevented), false);
        check(`${mode} ${key} does not navigate options for nonempty query`, await search.getAttribute('aria-activedescendant'), active);
      }
      await search.press('Meta+ArrowLeft'); await page.keyboard.type('X');
      check(`${mode} native caret shortcut edits query`, await search.inputValue(), 'XSite 1');
      await search.fill('Site 1'); await rows().first().waitFor();
      await search.press('ArrowDown'); await search.press('Enter');
      await page.waitForFunction(() => window.pickerFixture.confirmation.pending);
      check(`${mode} pending search retains focus`, await search.evaluate(node => node === document.activeElement), true);
      check(`${mode} pending search is readonly`, await search.getAttribute('readonly'), '');
      await settle('Selection rejected');
      check(`${mode} rejection retains popup and query`, [await search.count(), await search.inputValue()], [1, 'Site 1']);
      check(`${mode} rejection emits nothing`, await events(), '0');
      await rows().first().click(); await settle('Pointer rejected');
      check(`${mode} pointer rejection retains selection`, await selectedValues(subject), []);
      if (mode === 'single') {
        for (const key of ['Space', 'Alt+ArrowUp']) {
          await trigger.focus(); await trigger.press(key); await settle('Trigger rejected');
          check(`single ${key} rejection preserves popup`, await search.count(), 1);
          check(`single ${key} rejection emits nothing`, await events(), '0');
        }
      }
      await search.focus(); await search.press('ArrowDown'); await search.press('Enter'); await settle();
      await page.waitForFunction(id => JSON.parse(document.querySelector(`[data-picker-case="${id}"] [data-case-selection]`).textContent).length === 1, await subject.getAttribute('data-picker-case'));
      check(`${mode} accepted asynchronous event`, await events(), '1');
      check(`${mode} acceptance popup policy`, await search.count(), mode === 'single' ? 0 : 1);
      if (mode === 'single') {
        await open(); await page.locator('.fui-Option:visible[aria-selected="true"]').click();
        check('single selected reactivation emits nothing', await events(), '1');
        check('single selected reactivation keeps popup', await search.count(), 1);
      } else {
        await search.fill('Site 2'); await rows().filter({ hasText: 'Site 2' }).click(); await settle();
        check('multi native summary', await trigger.textContent(), '2 elementi selezionati');
        await page.locator('.fui-Option:visible[aria-checked="true"]').first().click(); await settle();
        check('multi selected row removes independently of query', (await selectedValues(subject)).length, 1);
      }
      await search.fill('Site'); await page.getByRole('button', { name: 'Cancella ricerca', exact: true }).click();
      check(`${mode} query clear only clears search`, await search.inputValue(), '');
      check(`${mode} query clear retains selection`, (await selectedValues(subject)).length, 1);
      const beforeIme = await events();
      await search.evaluate(node => {
        node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: 'あ' }));
        node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', isComposing: true, keyCode: 229 }));
        node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'あ' }));
      });
      check(`${mode} composition Enter does not select`, await events(), beforeIme);
      await search.press('Escape');
      check(`${mode} Escape restores trigger focus`, await trigger.evaluate(node => node === document.activeElement), true);
      await open(); await search.press('End'); await search.press('Tab'); await frames();
      check(`${mode} Tab closes without commit`, [await search.count(), await events()], [0, beforeIme]);
      check(`${mode} search Tab reaches next adjacent control`, await forwardTarget(), true);
      await open(); await search.press('Shift+Tab'); await frames();
      check(`${mode} Shift+Tab closes without commit`, [await search.count(), await events()], [0, beforeIme]);
      check(`${mode} search Shift+Tab reaches preceding control`, await backwardTarget(), true);
      await open(); await trigger.focus(); await trigger.press('Tab'); await frames();
      check(`${mode} trigger Tab never implicitly commits`, await events(), beforeIme);
      check(`${mode} trigger Tab reaches next adjacent control`, await forwardTarget(), true);
      await open(); await trigger.focus(); await trigger.press('Shift+Tab'); await frames();
      check(`${mode} trigger Shift+Tab reaches preceding control`, await backwardTarget(), true);
      await page.keyboard.press('Escape');
      await action(subject, 'defer').click(); await action(subject, 'blur').click();
      await open(); const beforeBlur = Number(await subject.locator('[data-case-validations]').textContent());
      await trigger.focus(); await search.focus(); await search.press('ArrowDown');
      check(`${mode} internal focus does not validate blur`, Number(await subject.locator('[data-case-validations]').textContent()), beforeBlur);
      await outside.click(); await frames();
      check(`${mode} outside exit validates once`, Number(await subject.locator('[data-case-validations]').textContent()), beforeBlur + 1);
      check(`${mode} outside exit preserves focus`, await outside.evaluate(node => node === document.activeElement), true);
      await action(subject, 'blur').click(); await action(subject, 'searchable').click();
      await trigger.click();
      check(`${mode} searchable false omits popup search`, await search.count(), 0);
      check(`${mode} searchable false keeps native options`, await rows().count() > 0, true);
      await page.keyboard.press('Escape'); await action(subject, 'searchable').click();
      if (mode === 'multi') {
        await action(subject, 'limit').click(); await trigger.click(); await rows().last().click(); await frames();
        check('multi selection limit retains existing value', (await selectedValues(subject)).length, 1);
        check('multi selection limit exposes validation', await subject.getByText('Il limite di selezione è stato raggiunto.', { exact: true }).count(), 1);
        await page.keyboard.press('Escape'); await action(subject, 'limit').click();
      }
      await action(subject, 'clear').click();
    }
    const delayed = card('SitePicker-combobox-single-narrow'), delayedTrigger = delayed.locator('.fui-Dropdown__button');
    await action(delayed, 'defer').click(); await delayedTrigger.click(); await query().fill('Site 3'); await rows().first().click();
    await page.waitForFunction(() => window.pickerFixture.confirmation.pending); await outside.click(); await settle();
    check('late acceptance preserves outside focus', await outside.evaluate(node => node === document.activeElement), true);
    check('late acceptance leaves dismissed popup closed', await page.locator('.fui-SearchBox input:visible').count(), 0);
    await action(delayed, 'defer').click(); await action(delayed, 'clear').click();
    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      for (const presentation of ['combobox', 'tagPicker']) for (const mode of ['single', 'multi']) {
        const subject = card(`${name}-${presentation}-${mode}-wide`), queryInput = subject.locator('.fui-Dropdown__button, .fui-TagPickerInput');
        await queryInput.click(); await query().fill('zzzznomatch');
        await page.getByRole('status').filter({ hasText: 'Nessun risultato per questa ricerca' }).waitFor();
        check(`${name}-${presentation}-${mode} completed empty query feedback`, await page.getByRole('status').last().textContent(), 'Nessun risultato per questa ricerca');
        await query().fill('x');
        check(`${name}-${presentation}-${mode} minimum guidance is the placeholder`, await query().getAttribute('placeholder'), 'Digita almeno 2 caratteri per cercare');
        check(`${name}-${presentation}-${mode} minimum guidance has no duplicate status`, await page.getByRole('status').filter({ hasText: 'Digita almeno 2 caratteri per cercare' }).count(), 0);
        await query().fill(''); await page.keyboard.press('Escape');
      }
    }

    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      const subject = card(`${name}-combobox-single-wide`), input = subject.locator('.fui-Dropdown__button');
      await input.click(); await rows().first().click();
      await page.waitForFunction(id => document.querySelector(`[data-picker-case="${id}"] [data-case-selection]`).textContent !== '[]', `${name}-combobox-single-wide`);
      const selected = await subject.locator('[data-case-selection]').textContent();
      await input.click(); await page.getByText('Selezionato', { exact: true }).waitFor();
      const checked = page.locator('.fui-Option:visible[aria-selected="true"]');
      check(`${name} current single row checked once`, await checked.count(), 1);
      check(`${name} current single check visible`, await checked.locator('.fui-Option__checkIcon').count(), 1);
      await query().fill('unrelated'); await page.getByText('Selezionato', { exact: true }).waitFor();
      check(`${name} single selection retained during unrelated search`, await checked.count(), 1);
      await checked.click();
      check(`${name} clicking current value preserves selection`, await subject.locator('[data-case-selection]').textContent(), selected);
      await subject.locator('.fui-Dropdown__clearButton').click();
      await page.waitForFunction(id => document.querySelector(`[data-picker-case="${id}"] [data-case-selection]`).textContent === '[]', `${name}-combobox-single-wide`);
      check(`${name} single clearable retained`, JSON.parse(await subject.locator('[data-case-selection]').textContent()), []);
      await page.keyboard.press('Escape');
    }
    for (const name of ['FieldPicker', 'ViewPicker', 'ContentTypePicker']) {
      const subject = card(`${name}-combobox-single-wide`), input = subject.locator('.fui-Dropdown__button');
      await input.click(); await rows().last().waitFor();
      check(`${name} complete catalog beyond thirty`, await rows().count(), 35);
      const before = await page.evaluate(() => window.pickerFixture.stats.sp);
      const title = name === 'FieldPicker' ? 'Field 34' : name === 'ContentTypePicker' ? 'Content type 34' : 'View 34';
      await query().fill(title.toUpperCase());
      await page.waitForFunction(() => [...document.querySelectorAll('.fui-Option')].filter(n => n.getBoundingClientRect().width).length === 1);
      check(`${name} local case insensitive match`, (await rows().first().textContent()).includes(title), true);
      await query().fill(title.toLowerCase());
      check(`${name} typing does not call REST`, await page.evaluate(() => window.pickerFixture.stats.sp), before);
      await page.keyboard.press('Escape');
    }
    const list = card('ListPicker-combobox-multi-wide'), listInput = list.locator('.fui-Dropdown__button');
    await action(list, 'empty').click(); await listInput.click(); await rows().first().waitFor();
    check('metadata catalog default thirty', await rows().count(), 30);
    check('catalog lookahead hint', (await page.getByRole('status').last().textContent()).includes('Sono disponibili altri elementi'), true);
    await query().fill('LIST 34'); await rows().first().waitFor();
    await page.waitForFunction(() => [...document.querySelectorAll('.fui-Option')].filter(n => n.getBoundingClientRect().width).length === 1);
    check('case insensitive search beyond initial thirty', (await rows().first().textContent()).includes('List 34'), true);
    await rows().first().click(); await page.waitForFunction(() => document.querySelector('[data-picker-case="ListPicker-combobox-multi-wide"] [data-case-selection]').textContent !== '[]');
    await query().fill('unrelated'); await page.getByText('Selezionati', { exact: true }).waitFor();
    check('selected remains checked with unrelated query', await rows().first().getAttribute('aria-checked'), 'true');
    await rows().first().click(); await page.waitForFunction(() => document.querySelector('[data-picker-case="ListPicker-combobox-multi-wide"] [data-case-selection]').textContent === '[]');
    check('multi Dropdown deselection without matching search', JSON.parse(await list.locator('[data-case-selection]').textContent()), []);
    await page.keyboard.press('Escape');
    const contextual = card('SitePicker-combobox-multi-wide'), contextInput = contextual.locator('.fui-Dropdown__button');
    await action(contextual, 'empty').click(); await action(contextual, 'context-sites').click(); await contextInput.click();
    await page.getByText('Sito corrente', { exact: true }).waitFor(); await page.getByText('Hub associato', { exact: true }).waitFor();
    check('Site contextual suggestions available without query', await rows().count(), 2);
    await query().fill('NO MATCH');
    await page.getByText('Sito corrente', { exact: true }).waitFor();
    await page.getByText('Hub associato', { exact: true }).waitFor();
    check('Site fixed choices survive unrelated typed search', await rows().count(), 2);
    await page.getByRole('status').filter({ hasText: 'Nessun risultato per questa ricerca' }).waitFor();
    check('fixed choices do not mask empty query feedback', await page.getByRole('status').last().textContent(), 'Nessun risultato per questa ricerca');
    await query().fill('Site 0');
    await page.getByRole('status').filter({ hasText: 'Nessun altro risultato per questa ricerca' }).waitFor();
    check('matching fixed site is not reported as zero matches', await page.getByRole('status').last().textContent(), 'Nessun altro risultato per questa ricerca');
    await query().fill('');
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
    const site = card('SitePicker-combobox-single-wide'), input = site.locator('.fui-Dropdown__button');
    await input.click(); await rows().first().click(); await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-single-wide"] [data-case-selection]').textContent !== '[]');
    const selection = await site.locator('[data-case-selection]').textContent(); check('site JSON reference selected', JSON.parse(selection)[0].kind, 'site');
    check('no separate clear-all button', await site.getByRole('button', { name: 'Svuota selezione', exact: true }).count(), 0);
    await site.locator('.fui-Dropdown__clearButton').click();
    await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-single-wide"] [data-case-selection]').textContent === '[]');
    check('native Dropdown clear emits empty references', JSON.parse(await site.locator('[data-case-selection]').textContent()), []);
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
    check('empty set guides search through its placeholder', await query().getAttribute('placeholder'), 'Digita almeno 2 caratteri per cercare');
    check('empty set has no duplicate minimum guidance', await page.getByRole('status').filter({ hasText: 'Digita almeno 2 caratteri per cercare' }).count(), 0);
    await query().fill('Site'); await rows().first().waitFor(); check('typed search includes nonconsumer sources', await rows().count(), 7);
    await page.keyboard.press('ArrowDown'); await page.keyboard.press('Enter');
    await page.waitForFunction(() => document.querySelector('[data-picker-case="SitePicker-combobox-single-wide"] [data-case-selection]').textContent !== '[]');
    check('keyboard selection confirms domain value', JSON.parse(await site.locator('[data-case-selection]').textContent())[0].kind, 'site');
    await page.keyboard.press('Escape');
    const paging = card('SitePicker-tagPicker-multi-narrow'), pagingInput = paging.locator('.fui-TagPickerInput');
    await action(paging, 'pages').click(); await pagingInput.click(); await pagingInput.fill('Site');
    await page.getByRole('status').filter({ hasText: 'Ricerca parziale: affina la ricerca.' }).waitFor();
    check('partial query shows bounded first page', await rows().count(), 2);
    await paging.getByRole('button', { name: 'Altri risultati', exact: true }).click();
    check('loadMore retains the query', await pagingInput.inputValue(), 'Site');
    check('loadMore never changes selection', await selectedValues(paging), []);
    if (await pagingInput.getAttribute('aria-expanded') !== 'true') await pagingInput.click();
    await page.waitForFunction(() => [...document.querySelectorAll('.fui-TagPickerOption')].filter(node => node.getBoundingClientRect().width).length === 4);
    check('loadMore accumulates the next actual source page', await rows().count(), 4);
    await pagingInput.fill(''); await page.keyboard.press('Escape'); await action(paging, 'pages').click();
    const principal = card('PrincipalPicker-tagPicker-multi-wide'), principalInput = principal.locator('.fui-TagPickerInput');
    await action(principal, 'empty').click(); await action(principal, 'personal').click(); await principalInput.click(); await rows().first().waitFor();
    await page.getByText('Persone suggerite', { exact: true }).waitFor(); await page.getByText('I tuoi gruppi', { exact: true }).waitFor();
    check('personal sources overall cap', await rows().count() <= 5, true); check('directory roles excluded', (await rows().allTextContents()).some(text => text.includes('directoryRole')), false);
    await rows().filter({ hasText: 'User 1' }).first().click();
    await page.waitForFunction(() => document.querySelector('[data-picker-case="PrincipalPicker-tagPicker-multi-wide"] [data-case-selection]').textContent !== '[]');
    if (await principalInput.getAttribute('aria-expanded') !== 'true') await principalInput.click();
    await rows().first().waitFor();
    check('TagPicker selected exclusion is observed in an open popup', await rows().count() > 0, true);
    check('TagPicker selected identity not duplicated in popup', await rows().filter({ hasText: 'User 1' }).count(), 0);
    check('TagPicker has no selected section', await page.getByText('Selezionati', { exact: true }).count(), 0);
    await page.screenshot({ path: path.join(evidence, 'principal-tagpicker-selected.png') });
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
    for (let index = 0; index < 56; index++) {
      const subject = cases.nth(index), id = await subject.getAttribute('data-picker-case');
      await action(subject, 'long').click();
      const longInput = subject.locator('.fui-Dropdown__button, .fui-TagPickerInput');
      await longInput.click(); await query().fill('');
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
      if (id.includes('combobox')) {
        await longRow.click(); await frames(); await page.keyboard.press('Escape');
        const trigger = subject.locator('.fui-Dropdown__button');
        const selectedLayout = await trigger.evaluate(node => {
          const root = node.closest('.fui-Dropdown'), chevron = node.querySelector('.fui-Dropdown__expandIcon'), media = node.querySelector('.fui-Avatar'), clear = root.querySelector('.fui-Dropdown__clearButton');
          const bounds = node.getBoundingClientRect(), chevronBounds = chevron.getBoundingClientRect();
          return { chevronDirect: chevron.parentElement === node, chevronVisible: chevronBounds.width > 0 && chevronBounds.right <= bounds.right + 1,
            media: media?.getBoundingClientRect().width, fits: root.scrollWidth <= root.clientWidth + 1,
            clearVisible: !clear || clear.getBoundingClientRect().width > 0 };
        });
        check(`${id} native chevron stays a direct sibling`, selectedLayout.chevronDirect, true);
        // Fluent swaps the selected-single chevron for its native clear button.
        check(`${id} native clear or chevron remains visible`, id.includes('single') ? selectedLayout.clearVisible : selectedLayout.chevronVisible, true);
        check(`${id} selected trigger stays within control`, selectedLayout.fits, true);
        check(`${id} compact selected media size`, selectedLayout.media, (await selectedValues(subject)).length === 1 ? id.endsWith('narrow') ? 24 : 20 : undefined);
        if (id === 'SitePicker-combobox-single-narrow' || id === 'SitePicker-combobox-multi-narrow') {
          await trigger.click(); await page.screenshot({ path: path.join(evidence, `${id}-selected.png`) });
        }
      }
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
    const themed = card('PrincipalPicker-combobox-single-narrow'), themedTrigger = themed.locator('.fui-Dropdown__button');
    check('themed production trigger retains its long selection', (await selectedValues(themed)).length, 1);
    for (const dark of [false, true]) {
      if (dark) await page.locator('[data-global-action="theme"]').click();
      for (const rtl of [false, true]) {
        if (rtl) await page.locator('[data-global-action="rtl"]').click();
        await themedTrigger.click(); await query().waitFor(); await frames();
        const surface = page.locator('.fui-Dropdown__listbox:visible');
        const geometry = await surface.evaluate(node => {
          const r = node.getBoundingClientRect(), list = node.querySelector('[role="listbox"], [role="menu"]');
          const content = node.querySelector('.fui-SearchBox').parentElement.getBoundingClientRect();
          const token = getComputedStyle(node).getPropertyValue('--colorNeutralBackground1').trim();
          const expected = /^#[0-9a-f]{6}$/i.test(token) ? `rgb(${parseInt(token.slice(1,3),16)}, ${parseInt(token.slice(3,5),16)}, ${parseInt(token.slice(5,7),16)})` : token;
          return { theme: getComputedStyle(node).backgroundColor === expected, direction: getComputedStyle(node).direction,
            bounded: content.height <= Math.min(320, innerHeight * .8) + 1, ownsList: !!list,
            viewport: r.left >= -1 && r.right <= innerWidth + 1 && r.top >= -1 && r.bottom <= innerHeight + 1 };
        });
        check(`theme ${dark} rtl ${rtl} popup inherits native tokens`, geometry.theme, true);
        check(`theme ${dark} rtl ${rtl} direction`, geometry.direction, rtl ? 'rtl' : 'ltr');
        check(`theme ${dark} rtl ${rtl} bounded native position`, geometry.bounded && geometry.viewport && geometry.ownsList, true);
        await page.screenshot({ path: path.join(evidence, `dropdown-theme-${dark ? 'dark' : 'light'}-${rtl ? 'rtl' : 'ltr'}.png`) });
        await page.keyboard.press('Escape');
        if (rtl) await page.locator('[data-global-action="rtl"]').click();
      }
      if (dark) await page.locator('[data-global-action="theme"]').click();
    }
    await page.setViewportSize({ width: 320, height: 400 }); await themedTrigger.click(); await query().waitFor(); await frames();
    check('320px viewport positions whole production popup within viewport', await page.locator('.fui-Dropdown__listbox:visible').evaluate(node => { const r = node.getBoundingClientRect(); return r.left >= -1 && r.right <= innerWidth + 1 && r.top >= -1 && r.bottom <= innerHeight + 1; }), true);
    await page.screenshot({ path: path.join(evidence, 'dropdown-320px.png') }); await page.keyboard.press('Escape');
    await page.setViewportSize({ width: 1280, height: 1000 });
    // Real native factories mount into a separate SPFx-owned root only after gallery acceptance.
    console.log(`Gallery acceptance before native fixture: ${checks.length} checks, ${layouts.length} layouts.`);
    const galleryCheckCount = checks.length + 1, nativeChecks = [], nativeReads = [], nativeGeometry = [];
    const nativeCheck = (name, actual, expected) => { assert.deepEqual(actual, expected, name); nativeChecks.push(name); };
    const native = page.locator('[data-native-picker-fixture]');
    const skeleton = native.locator('[data-property-pane-picker-control]');
    const overlay = native.locator('[data-auto-skeleton-overlay]');
    const nativeInput = native.locator('.fui-TagPickerInput');
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
    for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
      await page.evaluate(name => { window.pickerFixture.native.blockReads(); window.pickerFixture.native.mount(name, 'unresolved', 'combobox'); }, name);
      const nativeTrigger = native.locator('.fui-Dropdown__button');
      await nativeTrigger.waitFor({ state: 'attached' }); await overlay.waitFor();
      nativeCheck(`${name} native Dropdown uses rectangle geometry`, await overlay.locator('[data-auto-skeleton-shape="rectangle"]').count() > 0, true);
      nativeCheck(`${name} native Dropdown label remains visible`, await native.getByText(`Native ${name}`, { exact: true }).isVisible(), true);
      nativeCheck(`${name} native Dropdown hint remains visible`, await native.getByText('Native hydration hint', { exact: true }).isVisible(), true);
      await nativeTrigger.evaluate(node => { window.nativeRetainedInput = node; });
      await releaseNative();
      nativeCheck(`${name} native Dropdown keeps trigger DOM after hydration`, await nativeTrigger.evaluate(node => node === window.nativeRetainedInput), true);
      nativeCheck(`${name} native Dropdown displays hydrated value`, await nativeTrigger.locator(':scope > span > span').last().textContent(), ({ SitePicker: 'Site 0', ListPicker: 'List 0', FieldPicker: 'Field 0', ViewPicker: 'View 0', ContentTypePicker: 'Content type 0', ListItemPicker: 'Contract.pdf', PrincipalPicker: 'User 0' })[name]);
      const beforeBlur = await page.evaluate(() => window.pickerFixture.native.validationCalls.length);
      await nativeTrigger.click(); await query().waitFor(); await frames();
      nativeCheck(`${name} native portal search receives focus`, await query().evaluate(node => node === document.activeElement), true);
      await nativeTrigger.focus(); await query().focus(); await query().fill('zzzznomatch'); await frames();
      nativeCheck(`${name} native internal portal focus avoids blur validation`, await page.evaluate(() => window.pickerFixture.native.validationCalls.length), beforeBlur);
      await outside.click(); await frames();
      nativeCheck(`${name} native field validates once on outside exit`, await page.evaluate(() => window.pickerFixture.native.validationCalls.length), beforeBlur + 1);
      nativeCheck(`${name} native focus movements never write JSON`, await page.evaluate(() => window.pickerFixture.native.changes.length), 0);
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
    if (page) { const diagnostic = await page.evaluate(() => ({ stats: window.pickerFixture.stats, state: window.pickerFixture.state, active: document.activeElement?.outerHTML.slice(0, 700), open: [...document.querySelectorAll('[aria-expanded="true"]')].map(n => n.getAttribute('aria-label')), rows: [...document.querySelectorAll('[role="option"],[role="menuitemcheckbox"]')].filter(n => n.getBoundingClientRect().width).map(n => n.textContent), errors: document.querySelector('[data-picker-case="PrincipalPicker-tagPicker-multi-wide"] [data-case-errors]')?.textContent })); console.error(JSON.stringify(diagnostic)); await page.screenshot({ path: path.join(evidence, 'failure.png') }); }
    throw error;
  } finally { await browser?.close(); await new Promise(resolve => server.close(resolve)); }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
