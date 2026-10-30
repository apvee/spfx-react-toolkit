const path = require('node:path');

async function cardinality(panel, check, label, selector) {
  await check(`${label}: one section`, await panel.locator('[data-demo-section]').count(), 1);
  await check(`${label}: one configuration`, await panel.locator('[data-demo-configuration]').count(), 1);
  await check(`${label}: one subject`, await panel.locator('[data-demo-subject]').count(), 1);
  if (selector) await check(`${label}: one control`, await panel.locator(selector).count(), 1);
}
async function diagnostics(panel) {
  const toggle = panel.getByRole('button', { name: 'Advanced diagnostics', exact: true });
  if (await toggle.getAttribute('aria-expanded') !== 'true') await toggle.click();
}
async function formProduct(page, { url, check }) {
  await page.goto(`${url}/?form-only=1&product-demo=1`);
  const panel = page.locator('[data-form-demo="panel"]'); await panel.waitFor();
  await cardinality(panel, check, 'form initial', 'form'); await diagnostics(panel);
  await check('form initial hint explains report', await panel.locator('[data-form-demo=slot-hint]').textContent(), 'Summarize team results, decisions and next steps.');
  const select = name => panel.locator(`select[data-form-demo="${name}"]`);
  const live = async () => JSON.parse(await panel.locator('[data-form-demo="values"]').textContent());
  const reference = panel.locator('[name="reference"]'); await reference.fill('Draft conservato');
  const identity = await reference.elementHandle();
  await select('resize').selectOption('320'); await select('resize').selectOption('640');
  await check('form props preserve draft and DOM', await identity.evaluate(e => e === document.querySelector('[name="reference"]') && e.value === 'Draft conservato'), true); await identity.dispose();
  const adapters = ['Input','Textarea','SearchBox','Checkbox','Switch','SpinButton','Slider','Select','RadioGroup','Dropdown','DropdownMulti','CheckboxGroup','Combobox','ComboboxMulti','TagPicker','Rating','ColorPicker','SwatchPicker','DatePicker','TimePicker','DateTime','FileInput'];
  const fields = { Input:'custom', Textarea:'notes', SearchBox:'query', Checkbox:'approved', Switch:'enabled', SpinButton:'quantity', Slider:'progress', Select:'country', RadioGroup:'priority', Dropdown:'category', DropdownMulti:'categories', CheckboxGroup:'checks', Combobox:'combo', ComboboxMulti:'combos', TagPicker:'tags', Rating:'rating', ColorPicker:'accent', SwatchPicker:'swatch', DatePicker:'date', TimePicker:'time', DateTime:'appointment', FileInput:'files' };
  const expectedValues = {
    Input: 'Detail aggiornato', Textarea: 'Note aggiornata', SearchBox: 'Detail aggiornato',
    Checkbox: true, Switch: false, SpinButton: 7, Slider: 30, Select: 'pt', RadioGroup: 'it',
    Dropdown: 'pt', DropdownMulti: ['pt', 'it'], CheckboxGroup: ['it', 'pt'], Combobox: 'pt',
    ComboboxMulti: ['pt', 'it'], TagPicker: ['it', 'pt'], Rating: 5, SwatchPicker: 'red',
    DatePicker: { year: 2026, month: 10, day: 11, hour: 0, minute: 0 },
    TimePicker: { year: 1970, month: 1, day: 1, hour: 13, minute: 45 },
    DateTime: { date: { year: 2026, month: 10, day: 12, hour: 0, minute: 0 }, time: { year: 1970, month: 1, day: 1, hour: 14, minute: 30 } },
    FileInput: { count: 1, names: ['form-browser-selection.txt'] },
  };
  let submitted = 0;
  for (const adapter of adapters) {
    await select('adapter').selectOption(adapter); const control = panel.locator(`[data-form-control="${adapter}"]`);
    await cardinality(panel, check, `form ${adapter}`, 'form');
    await check(`form ${adapter} sole variable field`, await panel.locator('[data-form-control]').count(), 1);
    await check(`form ${adapter} common draft`, await reference.inputValue(), 'Draft conservato');
    const input = control.locator('input').first();
    if (adapter === 'Input' || adapter === 'SearchBox') { await input.fill('Detail aggiornato'); }
    if (adapter === 'Textarea') await control.locator('textarea').fill('Note aggiornata');
    if (adapter === 'Checkbox' || adapter === 'Switch') await input.setChecked(!(await input.isChecked()));
    if (adapter === 'SpinButton') { await input.fill('7'); await input.press('Tab'); }
    if (adapter === 'Slider') { await input.focus(); await input.press('ArrowRight'); }
    if (adapter === 'Select') await control.locator('select').selectOption('pt');
    if (adapter === 'RadioGroup') await control.getByRole('radio', { name: 'Italy', exact: true }).check();
    if (adapter === 'Dropdown' || adapter === 'DropdownMulti') { await control.getByRole('combobox').click(); await page.locator('.fui-Dropdown__listbox').getByRole(adapter === 'DropdownMulti' ? 'menuitemcheckbox' : 'option', { name: adapter === 'DropdownMulti' ? 'Italy' : 'Portugal', exact: true }).click(); await page.keyboard.press('Escape'); }
    if (adapter === 'CheckboxGroup') await control.getByRole('checkbox', { name: 'Portugal', exact: true }).check();
    if (adapter === 'Combobox' || adapter === 'ComboboxMulti' || adapter === 'TagPicker') { await input.fill(adapter === 'ComboboxMulti' ? 'Italy' : 'Port'); await input.press('ArrowDown'); await input.press('Enter'); await input.press('Tab'); }
    if (adapter === 'Rating') await control.getByRole('radio').last().check();
    if (adapter === 'ColorPicker') { const hue = control.getByRole('slider', { name: 'Hue', exact: true }); await hue.focus(); await hue.press('ArrowRight'); }
    if (adapter === 'SwatchPicker') await control.getByRole('radio', { name: 'Red', exact: true }).click();
    if (adapter === 'DatePicker' || adapter === 'TimePicker') { await input.fill(adapter === 'DatePicker' ? '2026-10-11' : '13:45'); await input.press('Tab'); }
    if (adapter === 'DateTime') { await control.getByLabel('Appointment date', { exact: true }).fill('2026-10-12'); await control.getByLabel('Appointment date', { exact: true }).press('Tab'); await control.getByRole('combobox', { name: 'Appointment time', exact: true }).fill('14:30'); await control.getByRole('combobox', { name: 'Appointment time', exact: true }).press('Tab'); }
    if (adapter === 'FileInput') await control.locator('input[type=file]').setInputFiles(path.resolve('tests/fixtures/form-browser-selection.txt'));
    const snapshot = (await live())[fields[adapter]];
    if (adapter === 'ColorPicker') {
      await check('form ColorPicker changes known initial hue', snapshot.h > 210, true);
      await check('form ColorPicker retains independent HSV channels', { s: snapshot.s, v: snapshot.v, a: snapshot.a }, { s: 0.7, v: 0.8, a: 1 });
    } else await check(`form ${adapter} independently expected native value`, snapshot, expectedValues[adapter]);
    await panel.locator('[data-form-demo="submit"]').click(); submitted++;
    await check(`form ${adapter} native submit callback`, (await panel.locator('[data-form-demo="status"]').textContent()).includes(`Saves: ${submitted}`), true);
    const saved = JSON.parse(await panel.locator('[data-form-demo="saved"]').textContent());
    await check(`form ${adapter} native value submitted`, saved[fields[adapter]], snapshot);
    // The required common field verifies that each adapter participates in the same validation/submit boundary.
    await reference.fill(''); await panel.locator('[data-form-demo="submit"]').click();
    await check(`form ${adapter} validation blocks callback`, (await panel.locator('[data-form-demo="status"]').textContent()).includes(`Saves: ${submitted}`), true);
    await reference.fill('Draft conservato');
    if (adapter === 'Input' || adapter === 'DatePicker' || adapter === 'TimePicker' || adapter === 'DateTime') {
      const invalidInput = adapter === 'DateTime' ? control.getByLabel('Appointment date', { exact: true }) : input;
      const invalidText = adapter === 'Input' ? '' : adapter === 'TimePicker' ? '99:99' : '2026-02-31';
      await invalidInput.fill(invalidText); await invalidInput.press('Tab'); await panel.locator('[data-form-demo="submit"]').click();
      await check(`form ${adapter} own invalid field blocks stale submit`, (await panel.locator('[data-form-demo="status"]').textContent()).includes(`Saves: ${submitted}`), true);
      await check(`form ${adapter} own invalid draft remains visible`, await invalidInput.inputValue(), invalidText);
      await invalidInput.fill(adapter === 'Input' ? 'Detail aggiornato' : adapter === 'TimePicker' ? '13:45' : adapter === 'DateTime' ? '2026-10-12' : '2026-10-11'); await invalidInput.press('Tab');
    }
  }
  await select('adapter').selectOption('Textarea'); await panel.locator('textarea').fill('Compatibile');
  await select('adapter').selectOption('Input'); await check('form compatible text adapter retains value', await panel.locator('[name="custom"]').inputValue(), 'Compatibile');
  await select('adapter').selectOption('SpinButton'); await check('form incompatible adapter explicit reset', (await panel.locator('[data-form-demo="status"]').textContent()).includes('new domain'), true);
  await check('form incompatible adapter restores actual quantity default', (await live()).quantity, 2);
  await panel.locator('[name="quantity"]').fill('9'); await panel.locator('[name="quantity"]').press('Tab');
  await select('adapter').selectOption('DatePicker'); await select('adapter').selectOption('SpinButton');
  await check('form return to edited adapter resets only target domain', (await live()).quantity, 2);
  await check('form reset retains independent text draft', (await live()).notes, 'Compatibile');
  await check('form reset retains common title draft', await reference.inputValue(), 'Draft conservato');
  for (const family of ['object','array','schema','dynamic','rules']) {
    await select('mode').selectOption(family); await cardinality(panel, check, `form ${family}`, 'form');
    if (family === 'object') { await panel.getByRole('textbox', { name: 'Owner', exact: true }).fill('Ada Rossi'); await check('form object writes nested model', (await live()).customer.name, 'Ada Rossi'); }
    if (family === 'array') { await panel.getByRole('textbox', { name: 'Activity name 1', exact: true }).fill('Analisi'); await check('form array writes nested model', (await live()).lines[0].code, 'Analisi'); }
    if (family === 'schema' || family === 'dynamic') {
      const activeInput = panel.getByRole('textbox', { name: family === 'schema' ? 'Quantity' : 'Title', exact: true });
      await check(`form ${family} pristine active dirty state`, (await panel.locator('[data-form-demo="status"]').textContent()).includes('Dirty: false'), true);
      await panel.getByRole('checkbox', { name: 'Disabled', exact: true }).check();
      await check(`form ${family} disables actual field`, await activeInput.isDisabled(), true);
      await panel.getByRole('checkbox', { name: 'Disabled', exact: true }).uncheck();
      await activeInput.fill(family === 'schema' ? '23' : 'Dinamico');
      await check(`form ${family} current dirty state`, (await panel.locator('[data-form-demo="status"]').textContent()).includes('Dirty: true'), true);
      await panel.locator('[data-form-demo="submit"]').click();
      await check(`form ${family} transformed actual output`, JSON.parse(await panel.locator('[data-form-demo="saved"]').textContent()), family === 'schema' ? { quantity: 23 } : { dynamic: 'Dinamico' });
      if (family === 'schema') {
        const countBefore = await panel.locator('[data-form-demo="status"]').textContent();
        await activeInput.fill('invalid'); await panel.locator('[data-form-demo="submit"]').click();
        await check('form schema invalid quantity blocks callback', (await panel.locator('[data-form-demo="status"]').textContent()).match(/Saves: \d+/)[0], countBefore.match(/Saves: \d+/)[0]);
        await check('form schema current invalid quantity error', (await panel.locator('[data-form-demo="schema-errors"]').textContent()).includes('Enter a finite number'), true);
      }
      await panel.locator('[data-form-demo="new-record"]').click();
      await check(`form ${family} load changes actual current model`, await activeInput.inputValue(), family === 'schema' ? '12' : 'Annual report');
      await check(`form ${family} load clears current dirty state`, (await panel.locator('[data-form-demo="status"]').textContent()).includes('Dirty: false'), true);
      await activeInput.fill(family === 'schema' ? '24' : 'Report pending');
      await panel.getByRole('checkbox', { name: 'Complete save manually', exact: true }).check();
      await panel.locator('[data-form-demo="submit"]').click();
      await check(`form ${family} pending locks family selector`, await select('mode').isDisabled(), true);
      await panel.locator('[data-form-demo="complete-save"]').click();
      await check(`form ${family} pending completed output`, JSON.parse(await panel.locator('[data-form-demo="saved"]').textContent()), family === 'schema' ? { quantity: 24 } : { dynamic: 'Report pending' });
      await panel.getByRole('checkbox', { name: 'Complete save manually', exact: true }).uncheck();
      await panel.getByRole('checkbox', { name: 'Simulate error', exact: true }).check(); await panel.locator('[data-form-demo="submit"]').click();
      await check(`form ${family} actual save error is visible`, (await panel.locator('[data-form-demo="status"]').textContent()).includes('Save failed'), true);
      await panel.getByRole('checkbox', { name: 'Simulate error', exact: true }).uncheck();
    }
  }
}
async function dataGridProduct(page, { url, check }) {
  await page.goto(`${url}/?data-grid-only=1&product-demo=1`);
  const panel = page.locator('#data-grid-demo'); await panel.waitFor(); await diagnostics(panel);
  const grid = panel.locator('#data-grid-main');
  // Toolkit status/error surfaces belong to the subject and may replace the native grid.
  const subject = panel.locator('[data-demo-subject]');
  for (const api of ['typed','descriptors','dynamic']) {
    await panel.locator('#dg-api').selectOption(api);
    for (const preset of api === 'typed' ? ['overview','dates','team','metrics','custom'] : ['overview']) {
      if (api === 'typed') await panel.locator('#dg-preset').selectOption(preset);
      await cardinality(panel, check, `grid ${api} ${preset}`, '[role=grid]');
      const columns = await grid.locator('.fui-DataGridHeader [data-grid-column-id]').count();
      await check(`grid ${api} ${preset} compact columns`, columns >= 4 && columns <= 5, true);
      if (api === 'typed') {
        const expected = { overview: ['title','total','status','owner','actions'], dates: ['title','enabled','day','instant'], team: ['title','link','owner','people','tags'], metrics: ['title','progress','rating','trend'], custom: ['title','enabled','icon','custom'] };
        await check(`grid ${preset} semantic columns`, await grid.locator('.fui-DataGridHeader [data-grid-column-id]').evaluateAll(nodes => nodes.map(e => e.getAttribute('data-grid-column-id'))), expected[preset]);
        if (preset === 'team') { await grid.locator('[data-grid-column-id=link] a').first().click(); await check('grid team actual navigation', JSON.parse(await panel.locator('#dg-navigation').textContent()).rowId, 'a'); }
        if (preset === 'metrics') { const prior = await grid.locator('[data-grid-column-id=trend]').last().textContent(); await panel.locator('#dg-update-sparkline').click(); await check('grid metrics updates real trend', (await grid.locator('[data-grid-column-id=trend]').last().textContent()) !== prior, true); }
        if (preset === 'custom') { await grid.locator('[data-custom-row=a]').click(); await check('grid custom cell owns live state', await grid.locator('[data-custom-row=a]').textContent(), 'Custom 1'); }
      }
    }
  }
  await check('grid dynamic reports active rows', (await panel.locator('[role=status]').first().textContent()).includes('Rows: 3'), true);
  await check('grid dynamic hides typed sort and trend commands', await panel.locator('#dg-apply-sort,#dg-update-sparkline,#dg-color').count(), 0);
  const rowIds = () => grid.locator('.fui-DataGridBody > [role=row]').evaluateAll(nodes => nodes.map(row => row.getAttribute('data-grid-row-id')));
  const rowsBefore = await rowIds(); await panel.locator('#dg-reverse').click(); await check('grid dynamic reverse changes actual supplied rows', await rowIds(), [...rowsBefore].reverse());
  await panel.locator('#dg-reverse').click();
  const mediumHeight = await grid.locator('.fui-DataGridBody > [role=row]').first().evaluate(e => e.getBoundingClientRect().height);
  await panel.locator('#dg-size').selectOption('small'); await check('grid dynamic size changes row geometry', await grid.locator('.fui-DataGridBody > [role=row]').first().evaluate(e => e.getBoundingClientRect().height) < mediumHeight, true);
  await panel.locator('#dg-selection-mode').selectOption('none'); await check('grid dynamic hides selection cells', await grid.getByRole('checkbox').count(), 0);
  const activationCell = grid.locator('.fui-DataGridBody > [role=row]').first().locator('[role=gridcell]').first();
  await activationCell.focus();
  await check('grid none mode keyboard starts at actual native cell', await activationCell.evaluate(cell => cell === document.activeElement), true);
  await page.keyboard.press('Enter');
  const activationOutput = panel.locator('#dg-activation-output');
  await activationOutput.filter({ hasText: '"rowId"' }).waitFor();
  await check('grid none mode enables real row activation', JSON.parse(await activationOutput.textContent()), { rowId: 'valid' });
  await panel.locator('#dg-selection-mode').selectOption('multiselect'); await grid.getByRole('checkbox').nth(1).check();
  await check('grid dynamic actual selected identity', JSON.parse(await panel.locator('#dg-selection').textContent()).includes('valid'), true);
  await panel.getByRole('checkbox', { name: 'Native resizing', exact: true }).check(); await check('grid dynamic resizing handles mounted', await grid.locator('.fui-TableResizeHandle').count() > 0, true); await panel.getByRole('checkbox', { name: 'Native resizing', exact: true }).uncheck();
  await panel.getByRole('checkbox', { name: 'Supply first row only', exact: true }).check(); await check('grid dynamic filtered row count', await grid.locator('.fui-DataGridBody > [role=row]').count(), 1); await panel.getByRole('checkbox', { name: 'Supply first row only', exact: true }).uncheck();
  await panel.getByRole('checkbox', { name: 'Hide all columns', exact: true }).check(); await check('grid dynamic all-hidden omits native grid', await grid.count(), 0); await panel.getByRole('checkbox', { name: 'Hide all columns', exact: true }).uncheck();
  await panel.getByRole('checkbox', { name: 'Empty rows', exact: true }).check(); await panel.locator('#dg-empty-state').selectOption('noResults'); await panel.getByRole('checkbox', { name: 'Custom empty and no-results slots', exact: true }).check();
  await check('grid dynamic custom no-results slot visible', await subject.getByRole('status').filter({ hasText: /^Consumer-owned no-results slot$/ }).isVisible(), true);
  await check('grid dynamic custom no-results slot', await subject.getByRole('status').filter({ hasText: /^Consumer-owned no-results slot$/ }).count(), 1);
  await panel.getByRole('checkbox', { name: 'Empty rows', exact: true }).uncheck(); await panel.getByRole('checkbox', { name: 'Custom empty and no-results slots', exact: true }).uncheck(); await panel.locator('#dg-empty-state').selectOption('empty');
  await panel.locator('#dg-api').selectOption('typed'); await panel.locator('#dg-preset').selectOption('overview');
  await grid.getByRole('checkbox').nth(1).check(); await check('grid selection actual intent', JSON.parse(await panel.locator('#dg-selection').textContent()).includes('a'), true);
  await grid.getByRole('button', { name: 'Open report', exact: true }).first().click(); await check('grid actual record action', JSON.parse(await panel.locator('#dg-action').textContent()), { rowId: 'a', actionId: 'review' });
  await panel.getByRole('checkbox', { name: 'Loading', exact: true }).check(); await cardinality(panel, check, 'grid pending', '[role=grid]'); await panel.getByRole('checkbox', { name: 'Loading', exact: true }).uncheck();
  await panel.getByRole('checkbox', { name: 'Error', exact: true }).check(); await check('grid error visible', await subject.getByText('Consumer-owned loading error', { exact: true }).count(), 1); await check('grid error message actually visible', await subject.getByText('Consumer-owned loading error', { exact: true }).isVisible(), true); await panel.getByRole('checkbox', { name: 'Error', exact: true }).uncheck();
  await panel.getByRole('checkbox', { name: 'Empty rows', exact: true }).check(); await check('grid empty actual rows', await grid.locator('.fui-DataGridBody > [role=row]').count(), 0);
}
async function stageProduct(page, { url, check }) {
  await page.goto(`${url}/?product-demo=1`); const panel = page.locator('[data-stage-demo="panel"]'); await panel.waitFor(); await diagnostics(panel);
  const waitForClosedFocus = async () => {
    await panel.locator('[data-stage-demo="open"]').filter({ hasText: 'Open: false' }).waitFor();
    // Native Fluent focus restoration follows surface exit motion/removal.
    // Match the legacy driver's paint/motion boundary without assigning focus.
    await page.evaluate(async () => {
      const paint = () => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      await paint();
      await Promise.all(document.getAnimations().filter(animation => animation.effect?.getTiming().iterations !== Infinity).map(animation => animation.finished.catch(() => undefined)));
      await paint();
    });
    await page.waitForFunction(() => {
      const trigger = document.querySelector('[data-stage-demo="panel"] [data-stage-demo="trigger"]');
      return trigger !== null && trigger === document.activeElement;
    });
  };
  for (const mode of ['dialog','drawer','popover']) {
    await panel.locator('select[data-stage-demo="mode"]').selectOption(mode); await cardinality(panel, check, `stage ${mode}`);
    await panel.locator('[data-stage-demo="trigger"]').click(); const surface = page.locator('[data-stage-demo="surface"]:visible'); await surface.waitFor();
    await check(`stage ${mode} one form`, await surface.locator('form').count(), 1); await check(`stage ${mode} two fields`, await surface.locator('input').count(), 2);
    await check(`stage ${mode} presentation locked`, await panel.locator('select[data-stage-demo="mode"]').isDisabled(), true);
    await surface.locator('[name="name"]').fill(`Draft ${mode}`); await surface.locator('[data-stage-demo="save"]').click();
    await check(`stage ${mode} actual save`, (await panel.locator('[data-stage-demo="status"]').textContent()).includes(`Saved Draft ${mode}`), true);
    await surface.getByRole('button', { name: 'Close A', exact: true }).click();
    await waitForClosedFocus();
    await check(`stage ${mode} focus returned`, await panel.locator('[data-stage-demo="trigger"]').evaluate(e => e === document.activeElement), true);
    await panel.getByRole('checkbox', { name: 'Complete save manually', exact: true }).check();
    await panel.locator('[data-stage-demo="trigger"]').click(); await surface.waitFor();
    await surface.locator('[name="name"]').fill(`Pending ${mode}`); await surface.locator('[data-stage-demo="save"]').click();
    await surface.locator('[data-stage-demo="pending"]').filter({ hasText: 'Pending: true' }).waitFor();
    await surface.getByRole('button', { name: 'Close A', exact: true }).click();
    await check(`stage ${mode} pending dismissal rejected`, (await panel.locator('[data-stage-demo="status"]').textContent()).includes('Close refused: save pending'), true);
    await check(`stage ${mode} pending guard retains surface`, await surface.count(), 1);
    await check(`stage ${mode} pending guard retains focus in surface`, await surface.evaluate(e => e.contains(document.activeElement)), true);
    await surface.locator('[data-stage-demo="complete-save"]').click(); await surface.locator('[data-stage-demo="pending"]').filter({ hasText: 'Pending: false' }).waitFor();
    await check(`stage ${mode} manual completion saves compact draft`, (await panel.locator('[data-stage-demo="status"]').textContent()).includes(`Saved Pending ${mode}`), true);
    await surface.getByRole('button', { name: 'Close A', exact: true }).click();
    await waitForClosedFocus();
    await check(`stage ${mode} completed session returns focus`, await panel.locator('[data-stage-demo="trigger"]').evaluate(e => e === document.activeElement), true);
    await panel.getByRole('checkbox', { name: 'Complete save manually', exact: true }).uncheck();
  }
  await panel.locator('[data-stage-demo="trigger"]').click(); const surface = page.locator('[data-stage-demo="surface"]:visible'); await surface.locator('[name="name"]').fill('Non salvato');
  await surface.getByRole('button', { name: 'Close A', exact: true }).click(); await check('stage dirty guard stays open', await surface.count(), 1);
  await surface.locator('[data-stage-demo="discard"]').click(); await check('stage discard closes', await page.locator('[data-stage-demo="surface"]:visible').count(), 0);
}
async function descriptionListProduct(page, { url, check }) {
  await page.goto(`${url}/?product-demo=1`); const panel = page.locator('[data-description-list-demo="panel"]'); await panel.waitFor(); await diagnostics(panel);
  for (const api of ['data','compound']) {
    await panel.locator('select[data-demo-variant="API"]').selectOption(api); await cardinality(panel, check, `details ${api}`, 'dl');
    await check(`details ${api} realistic five pairs`, await panel.locator('dt').count(), 5);
    const counter = panel.locator('[data-description-list-demo="counter"]');
    const initialCount = Number((await counter.textContent()).split(': ').at(-1));
    await counter.click(); await panel.getByRole('checkbox', { name: 'Reverse stable IDs', exact: true }).check();
    await panel.locator('select[data-description-list-demo="layout"]').selectOption('vertical');
    await check(`details ${api} reorder retains state`, await counter.textContent(), `Revisions: ${initialCount + 1}`);
    await panel.getByRole('checkbox', { name: 'Reverse stable IDs', exact: true }).uncheck();
  }
  await panel.getByRole('checkbox', { name: 'Empty input', exact: true }).check(); await cardinality(panel, check, 'details empty', 'dl'); await check('details empty pairs', await panel.locator('dt').count(), 0);
}
async function paginationProduct(page, { url, check }) {
  await page.goto(`${url}/?product-demo=1`); const panel = page.locator('[data-pagination-demo="panel"]'); await panel.waitFor(); await diagnostics(panel);
  for (const mode of ['known','unknown','cursor','loadMore']) {
    await panel.locator('select[data-demo-variant="Mode"]').selectOption(mode); await cardinality(panel, check, `pagination ${mode}`, 'nav');
    await check(`pagination ${mode} five visible results`, await panel.locator('[data-pagination-demo="results"] li').count(), 5);
    await panel.locator('[data-pagination-demo="navigation"]').getByRole('button', { name: mode === 'loadMore' ? 'Load additional orders' : 'Next page', exact: true }).click();
    await check(`pagination ${mode} actual next results`, await panel.locator('[data-pagination-demo="results"] li').first().textContent(), mode === 'loadMore' ? 'Order 01 · Hardware' : 'Order 06 · Software');
  }
  await panel.getByRole('checkbox', { name: 'Complete requests manually', exact: true }).check();
  await panel.locator('[data-pagination-demo="navigation"]').getByRole('button', { name: 'Load additional orders', exact: true }).click();
  await check('pagination manual request busy', await panel.locator('nav').getAttribute('aria-busy'), 'true');
  await panel.locator('[data-pagination-demo="complete"]').click(); await check('pagination completion adds real results', await panel.locator('[data-pagination-demo="results"] li').count(), 15);
  await panel.locator('select[data-pagination-demo="filter"]').selectOption('missing'); await check('pagination empty results', await panel.locator('[data-pagination-demo="results"] li').count(), 0);
}
async function groupsProduct(page, { check }) {
  const panel = page.locator('[data-groups-demo="panel"]'); await panel.waitFor(); await diagnostics(panel);
  const rows = panel.locator('[data-groups-demo=settings] [data-settings-group-item]');
  await check('groups settings preserves two separate original rows', await rows.count(), 2);
  await check('groups delivery row contains only email input', await rows.nth(0).locator('input').count(), 1);
  await check('groups delivery row has no notification switch', await rows.nth(0).getByRole('switch').count(), 0);
  await check('groups alerts row contains one switch', await rows.nth(1).getByRole('switch').count(), 1);
  await check('groups settings row icons are real SVGs', await rows.locator('[data-group-item-icon] svg').count(), 2);
  await check('groups original delivery caption', (await rows.nth(0).textContent()).trim(), 'Delivery addressUse your work email.');
  await check('groups original alert caption', (await rows.nth(1).textContent()).trim(), 'Account alertsReceive notifications about your account.');
  const kind = panel.locator('select[data-demo-variant="GroupType"]'); const email = panel.locator('input[data-groups-demo="email"]');
  await email.fill('draft@example.com'); await panel.getByRole('switch', { name: 'Receive account alerts', exact: true }).uncheck();
  await panel.locator('[data-groups-demo="action"]').click(); await check('groups actual action', (await panel.locator('[data-groups-demo="status"]').textContent()).includes('Actions: 1'), true);
  await kind.selectOption('accordion'); await cardinality(panel, check, 'groups accordion', '[data-groups-demo=accordion]');
  await check('groups accordion excludes delivery controls', await email.count(), 0);
  await check('groups accordion item icons are real SVGs', await panel.locator('[data-groups-demo=accordion] [data-group-item-icon] svg').count(), 2);
  const profile = panel.getByRole('button', { name: /^Profile details\s*Your account preferences\.$/ });
  const advanced = panel.getByRole('button', { name: 'Advanced preferences', exact: true });
  const note = panel.locator('input[data-groups-demo="note"]'); await note.fill('Retained draft');
  await profile.click(); await profile.click(); await check('groups close retains note', await note.inputValue(), 'Retained draft');
  await panel.getByRole('checkbox', { name: 'Allow multiple open sections', exact: true }).check(); await advanced.click();
  await check('groups multiple retains open profile', await profile.getAttribute('aria-expanded'), 'true');
  await check('groups multiple opens advanced panel', await advanced.getAttribute('aria-expanded'), 'true');
  const alerts = panel.getByRole('switch', { name: 'Enable advanced notifications', exact: true });
  await check('groups shared notification state', await alerts.isChecked(), false); await alerts.check();
  await panel.getByRole('checkbox', { name: 'Unmount closed sections (resets the note)', exact: true }).check();
  await profile.click();
  await note.waitFor({ state: 'detached' });
  await check('groups close actually unmounts note', await note.count(), 0);
  await profile.click();
  await note.waitFor({ state: 'visible' });
  await check('groups unmount resets note', await note.inputValue(), 'Retained note');
  await kind.selectOption('settings'); await cardinality(panel, check, 'groups settings', '[data-groups-demo=settings]'); await check('groups return retains shared draft', await email.inputValue(), 'draft@example.com');
  await check('groups return retains notification change', await panel.getByRole('switch', { name: 'Receive account alerts', exact: true }).isChecked(), true);
}
module.exports = { formProduct, dataGridProduct, stageProduct, descriptionListProduct, paginationProduct, groupsProduct };
