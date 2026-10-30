const { dataGridProduct } = require('../tests/fixtures/demo-control-isolation/task3-product.browser.cjs');
const path = require('node:path');
const fs = require('node:fs');
const os = require('node:os');

/** Native Fluent interactions against the same sample in source and compiled builds. */
async function exerciseDataGrid({ page, url, output, label, boundary, check, consoleMessages }) {
  const node = id => page.locator(`#${id}`);
  const main = node('data-grid-main');
  const rows = () => main.locator('.fui-DataGridBody > [role=row]');
  const json = async id => JSON.parse(await node(id).textContent());
  const selectedIds = async () => (await json('dg-selection')).sort();
  const settle = () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  const rowNames = () => rows().evaluateAll(elements => elements.map(row => row.querySelector('[data-grid-column-id=title]')?.textContent));
  const capture = async name => { await node('data-grid-demo').screenshot({ path: path.join(output, `${label}-data-grid-${name}.png`) }); };
  async function alignment(name) {
    await check(`DataGrid ${name} header/body alignment`, await main.evaluate(grid => {
      const heads = [...grid.querySelectorAll('.fui-DataGridHeader [data-grid-column-id]')];
      const first = grid.querySelector('.fui-DataGridBody [role=row]');
      const cells = first ? [...first.querySelectorAll(':scope > [data-grid-column-id]')] : [];
      return heads.length === 16 && cells.length === 16 && heads.every((head, i) => {
        if (head.getAttribute('data-grid-column-id') !== cells[i].getAttribute('data-grid-column-id')) return false;
        const h = head.getBoundingClientRect(), c = cells[i].getBoundingClientRect();
        return Math.abs(h.width - c.width) <= 1 && Math.abs(h.left - c.left) <= 1;
      });
    }), true);
  }
  async function actionLabelGeometry(name) {
    const inline = rows().first().getByRole('button', { name: 'Review this localized record in detail', exact: true });
    await check(`DataGrid ${name} long inline action label fits Actions cell`, await inline.evaluate(button => {
      const cell = button.closest('[role=gridcell]');
      const cellRect = cell.getBoundingClientRect(), buttonRect = button.getBoundingClientRect();
      const text = document.createRange(); text.selectNodeContents(button);
      const inside = rect => rect.left >= cellRect.left - 1 && rect.right <= cellRect.right + 1 && rect.top >= cellRect.top - 1 && rect.bottom <= cellRect.bottom + 1;
      return button.textContent === 'Review this localized record in detail' && inside(buttonRect) && [...text.getClientRects()].every(inside);
    }), true);
  }
  const loadingKinds = { title: 'text', total: 'number', enabled: 'boolean', day: 'date', instant: 'dateTime', status: 'badge', icon: 'icon', link: 'link', owner: 'person', people: 'people', tags: 'tags', progress: 'progress', rating: 'rating', trend: 'sparkline', actions: 'actions', custom: 'custom' };
  async function loadingGeometry() {
    return main.evaluate(grid => {
      const rect = element => {
        const box = element.getBoundingClientRect();
        return { left: box.left, top: box.top, right: box.right, bottom: box.bottom, width: box.width, height: box.height };
      };
      const corners = element => {
        const style = getComputedStyle(element);
        return [style.borderTopLeftRadius, style.borderTopRightRadius, style.borderBottomLeftRadius, style.borderBottomRightRadius].map(parseFloat);
      };
      return {
        cells: [...grid.querySelectorAll('.fui-DataGridBody [data-grid-column-id]')].map(cell => {
          const template = cell.querySelector('[data-grid-loading-kind]');
          const style = getComputedStyle(cell), box = rect(cell);
          return {
            column: cell.getAttribute('data-grid-column-id'), bounds: box,
            contentLeft: box.left + parseFloat(style.paddingLeft), contentRight: box.right - parseFloat(style.paddingRight),
            templateCount: cell.querySelectorAll('[data-grid-loading-kind]').length,
            kind: template?.getAttribute('data-grid-loading-kind'), row: template?.getAttribute('data-grid-loading-row-id'),
            templateColumn: template?.getAttribute('data-grid-loading-column-id'),
            primitives: template ? [...template.querySelectorAll('[data-skeleton-shape]')].map(shape => ({
              rect: rect(shape), shape: shape.getAttribute('data-skeleton-shape'), radius: corners(shape),
              medium: parseFloat(getComputedStyle(shape).getPropertyValue('--borderRadiusMedium')),
              className: shape.className, text: shape.textContent,
            })) : [],
          };
        }),
        bones: [...grid.querySelectorAll('[data-auto-skeleton-overlay] [data-auto-skeleton-shape]')].map(bone => ({ rect: rect(bone), shape: bone.getAttribute('data-auto-skeleton-shape'), radius: corners(bone) })),
      };
    });
  }
  const loadingFingerprint = geometry => geometry.cells.map(cell => ({ column: cell.column, kind: cell.kind, row: cell.row,
    primitives: cell.primitives.map(primitive => ({ className: primitive.className, text: primitive.text, shape: primitive.shape })) }));
  function loadingBoundsMatch(geometry) {
    const primitives = geometry.cells.flatMap(cell => cell.primitives);
    return primitives.length > 0 && geometry.bones.length === primitives.length && geometry.bones.every(bone => bone.shape === 'rectangle') && geometry.cells.every(cell =>
      cell.primitives.every((primitive, index) => {
        const box = primitive.rect;
        const bounded = box.width >= 8 && box.height >= 8 && box.left >= cell.contentLeft - 0.1 && box.right <= cell.contentRight + 0.1 && box.top >= cell.bounds.top - 0.1 && box.bottom <= cell.bounds.bottom + 0.1;
        const separated = cell.primitives.slice(index + 1).every(other => Math.min(box.right, other.rect.right) - Math.max(box.left, other.rect.left) <= 0.1 || Math.min(box.bottom, other.rect.bottom) - Math.max(box.top, other.rect.top) <= 0.1);
        const radius = Math.min(primitive.medium, Math.round(box.width) / 2, Math.round(box.height) / 2);
        return bounded && separated && primitive.shape === 'rectangle' && primitive.medium > 0 && primitive.radius.every(corner => corner === primitive.medium) && geometry.bones.some(bone =>
          ['left', 'top', 'width', 'height'].every(property => Math.abs(bone.rect[property] - box[property]) <= 1) && bone.radius.every(corner => corner === radius));
      }));
  }
  async function loadingCellSkeletons(name) {
    const geometry = await loadingGeometry();
    await check(`DataGrid ${name} every visible cell has a kind and identity-aware loading template`, geometry.cells.length === 64 && geometry.cells.every(cell => cell.templateCount === 1 && cell.kind === loadingKinds[cell.column] && cell.templateColumn === cell.column && typeof cell.row === 'string' && cell.row.length > 0), true);
    await check(`DataGrid ${name} templates follow each column's value geometry`, geometry.cells.length === 64 && geometry.cells.every(cell => {
      const shapes = cell.primitives, kind = cell.kind;
      if (kind === 'people' || kind === 'tags') return shapes.length === 1 && shapes[0].rect.height === 24;
      if (kind === 'person') return shapes.length === 2 && shapes[0].rect.height === 32 && shapes[1].rect.height === 16;
      if (kind === 'rating') return shapes.length === 5 && shapes.every(shape => shape.rect.width === 20 && shape.rect.height === 20);
      if (kind === 'sparkline') return shapes.length === 2 && shapes[0].rect.height === 28 && shapes[1].rect.height === 16;
      if (kind === 'actions') return shapes.length === 2 && shapes[0].rect.height === 20 && shapes[1].rect.height === 16;
      if (kind === 'progress') return shapes.length === 2 && shapes[0].rect.height === 8 && shapes[1].rect.height === 16;
      return shapes.length === 1 && shapes[0].rect.height === (kind === 'boolean' || kind === 'icon' ? 20 : kind === 'badge' ? 24 : 16);
    }), true);
    await check(`DataGrid ${name} rounded rectangle bones match bounded non-overlapping template geometry`, loadingBoundsMatch(geometry), true);
    const variedColumns = Object.keys(loadingKinds).filter(column => ['text', 'number', 'date', 'dateTime', 'link', 'custom'].includes(loadingKinds[column])).filter(column => new Set(geometry.cells.filter(cell => cell.column === column).map(cell => cell.primitives[0]?.rect.width)).size > 1);
    await check(`DataGrid ${name} simulated scalar widths vary across row identities`, variedColumns.length > 0, true);
    fs.writeFileSync(path.join(output, `${label}-data-grid-${name.replaceAll(' ', '-')}-cell-geometry.json`), JSON.stringify({ ...geometry, variedColumns }, null, 2));
    return geometry;
  }
  page.setDefaultTimeout(10000);
  await dataGridProduct(page, { url, check });
  await page.goto(`${url}/?data-grid-only=1`);
  await node('data-grid-demo').waitFor();
  const diagnostics = node('data-grid-demo').getByRole('button', { name: 'Advanced diagnostics', exact: true });
  if (await diagnostics.getAttribute('aria-expanded') !== 'true') await diagnostics.click();
  await page.getByRole('button', { name: 'Descriptors, dynamic values and row activation', exact: true }).click();
  await check('DataGrid local origin', new URL(page.url()).origin, new URL(url).origin);
  await check('DataGrid page identity', await page.title(), boundary === 'compiled' ? 'StylesPanel local browser fixture' : 'useSx real-browser checkpoint');
  await check('DataGrid explicit library boundary', await page.locator('[data-grid-boundary]').getAttribute('data-grid-boundary'), boundary);
  await check('DataGrid framework overlay absent', await page.locator('webpack-dev-server-client-overlay,#webpack-dev-server-client-overlay,nextjs-portal').count(), 0);
  await check('DataGrid native role', await main.getAttribute('role'), 'grid');
  await check('DataGrid full sixteen-column sample', await main.locator('.fui-DataGridHeader [role=columnheader][data-grid-column-id]').count(), 16);
  const mediumHeight = await rows().first().evaluate(row => row.getBoundingClientRect().height);
  await node('dg-size').selectOption('small');
  await check('DataGrid size changes native grid presentation', (await rows().first().evaluate(row => row.getBoundingClientRect().height)) < mediumHeight, true);
  await node('dg-size').selectOption('medium');
  await node('dg-selection-mode').selectOption('none');
  await check('DataGrid selectionMode none suppresses selection cells', await main.locator('.fui-DataGridSelectionCell').count(), 0);
  await check('DataGrid switching selection mode retains consumer IDs', await selectedIds(), ['off-page']);
  await node('dg-selection-mode').selectOption('multiselect');
  await check('DataGrid native header selection gridcell', await main.locator('.fui-DataGridHeader .fui-DataGridSelectionCell[role=gridcell]').count(), 1);
  await check('DataGrid native selection header/body alignment', await main.evaluate(grid => {
    const head = grid.querySelector('.fui-DataGridHeader .fui-DataGridSelectionCell').getBoundingClientRect();
    const body = grid.querySelector('.fui-DataGridBody .fui-DataGridSelectionCell').getBoundingClientRect();
    return Math.abs(head.width - body.width) <= 1 && Math.abs(head.left - body.left) <= 1;
  }), true);
  await check('DataGrid typed descriptor scenario', await node('data-grid-descriptors').getAttribute('role'), 'grid');
  await check('DataGrid dynamic scenario', await node('data-grid-dynamic').getAttribute('role'), 'grid');
  await node('dg-inspect-ref').click();
  await check('DataGrid ref native element', await node('dg-ref').textContent(), 'div');
  await node('dg-inspect-slot-refs').click();
  await check('DataGrid actual header and missing-value inner slot refs', await node('dg-slot-refs').textContent(), 'header=SPAN error=absent missing=SPAN');
  await node('dg-slot-suppress').check();
  await node('dg-inspect-slot-refs').click();
  await check('DataGrid false placeholder clears inner slot ref', await node('dg-slot-refs').textContent(), 'header=SPAN error=absent missing=absent');
  await node('dg-slot-suppress').uncheck();

  await alignment('intrinsic');
  const initialNames = await rowNames();
  await check('DataGrid supplied row names', initialNames, ['Alpha', 'Beta']);
  const booleanCells = main.locator('.fui-DataGridBody [data-grid-column-id=enabled]');
  await check('DataGrid boolean native icons measure 20px', await booleanCells.evaluateAll(cells => cells.every(cell => {
    const rect = cell.querySelector('svg').getBoundingClientRect();
    return rect.width === 20 && rect.height === 20;
  })), true);
  await check('DataGrid boolean native icons vertically centered', await booleanCells.evaluateAll(cells => cells.every(cell => {
    const rect = cell.getBoundingClientRect(), icon = cell.querySelector('svg').getBoundingClientRect();
    return Math.abs((rect.top + rect.bottom - icon.top - icon.bottom) / 2) <= 1;
  })), true);
  await check('DataGrid boolean accessible labels preserved', await booleanCells.locator('[role=img]').evaluateAll(elements => elements.map(element => element.getAttribute('aria-label'))), ['Enabled yes', 'Enabled no']);
  await booleanCells.first().scrollIntoViewIfNeeded();
  await page.screenshot({ path: path.join(output, `${label}-data-grid-boolean-icons.png`) });
  const nativeIconsMeasure20 = icons => icons.evaluateAll(elements => elements.length > 0 && elements.every(icon => {
    const rect = icon.getBoundingClientRect();
    return rect.width === 20 && rect.height === 20;
  }));
  await check('DataGrid Icon sample native icons measure 20px', await nativeIconsMeasure20(main.locator('.fui-DataGridBody [data-grid-column-id=icon] svg')), true);
  await check('DataGrid action sample native icons measure 20px', await nativeIconsMeasure20(rows().first().getByRole('button', { name: 'Disabled action', exact: true }).locator('svg')), true);
  await check('DataGrid More actions native icons measure 20px', await nativeIconsMeasure20(main.getByRole('button', { name: /more actions/i }).locator('svg')), true);
  const ratings = main.locator('.fui-DataGridBody [data-grid-column-id=rating] .fui-RatingDisplay');
  await check('DataGrid rating native star icons measure 20px', await nativeIconsMeasure20(ratings.locator('svg')), true);
  await check('DataGrid rating omits visible numeric value', await ratings.locator('.fui-RatingDisplay__valueText').count(), 0);
  await check('DataGrid rating omits visible vote count', await ratings.locator('.fui-RatingDisplay__countText').count(), 0);
  await check('DataGrid rating accessible name preserves value and scale', await main.getByRole('img', { name: '4/5', exact: true }).count(), 1);
  await check('DataGrid rating selected stars use native marigold token', await ratings.locator('.fui-RatingItem__selectedIcon svg').evaluateAll(icons => icons.length > 0 && icons.every(icon => {
    const token = getComputedStyle(icon).getPropertyValue('--colorPaletteMarigoldBorderActive').trim();
    const expected = /^#[\da-f]{6}$/i.test(token) ? `rgb(${[1, 3, 5].map(index => parseInt(token.slice(index, index + 2), 16)).join(', ')})` : token;
    return expected.length > 0 && getComputedStyle(icon).color === expected;
  })), true);

  await actionLabelGeometry('intrinsic');
  const sortable = main.locator('.fui-DataGridHeader [role=columnheader]').filter({ hasText: 'Title' }).first();
  await sortable.click(); await settle();
  await check('DataGrid pointer sort emits native ascending intent', await json('dg-sort'), { columnId: 'title', direction: 'ascending' });
  await check('DataGrid pointer sort native header state', await sortable.getAttribute('aria-sort'), 'ascending');
  await check('DataGrid sort intent preserves supplied order', await rowNames(), initialNames);
  await sortable.focus(); await page.keyboard.press('Enter'); await settle();
  await check('DataGrid keyboard sort emits native descending intent', await json('dg-sort'), { columnId: 'title', direction: 'descending' });
  await check('DataGrid keyboard sort native header state', await sortable.getAttribute('aria-sort'), 'descending');
  await check('DataGrid keyboard sort intent preserves supplied order', await rowNames(), initialNames);
  await node('dg-apply-sort').click(); await settle();
  await check('DataGrid consumer explicitly applies descending sort', await rowNames(), ['Beta', 'Alpha']);
  await node('dg-reverse').click(); await settle();
  await check('DataGrid independent consumer reverse restores initial order', await rowNames(), ['Alpha', 'Beta']);
  await node('dg-reverse').click(); await settle();
  await check('DataGrid consumer alone changes supplied order', await rowNames(), [...initialNames].reverse());
  await node('dg-reverse').click();
  const headerSlotClicks = Number(await node('dg-slot-clicks').textContent());
  await main.locator('[data-grid-slot-callback=header]').click();
  await check('DataGrid native header callback retains click handler', Number(await node('dg-slot-clicks').textContent()), headerSlotClicks + 1);
  await check('DataGrid initial selection exact IDs', await selectedIds(), ['off-page']);
  const headerSelection = main.locator('.fui-DataGridHeader').getByRole('checkbox').first();
  await headerSelection.click(); await settle();
  await check('DataGrid select-all selects supplied rows', await rows().evaluateAll(elements => elements.every(row => row.getAttribute('aria-selected') === 'true')), true);
  await check('DataGrid select-all exact IDs retains off-page', await selectedIds(), ['a', 'b', 'off-page']);
  const selected = ['a', 'b', 'off-page'];
  await node('dg-page').check(); await settle();
  await check('DataGrid removed supplied row keeps selected identity', await selectedIds(), selected);
  await headerSelection.click(); await settle();
  await check('DataGrid deselect supplied page exact retained IDs', await selectedIds(), ['b', 'off-page']);
  await node('dg-page').uncheck();
  const firstSelection = rows().first().getByRole('checkbox').first();
  await firstSelection.focus(); await page.keyboard.press('Space');
  await check('DataGrid native keyboard selection exact IDs', await selectedIds(), ['a', 'b', 'off-page']);
  const firstCell = rows().first().locator('[role=gridcell]').nth(1);
  await firstCell.focus(); await page.keyboard.press('ArrowDown');
  await check('DataGrid real keyboard row navigation', await rows().nth(1).evaluate(row => row.contains(document.activeElement)), true);
  await page.keyboard.press('ArrowRight');
  await check('DataGrid real keyboard column navigation', await rows().nth(1).evaluate(row => document.activeElement !== row.querySelectorAll('[role=gridcell]')[1] && row.contains(document.activeElement)), true);
  const activation = node('data-grid-activation');
  const activationRow = activation.locator('.fui-DataGridBody > [role=row]').first();
  const plainActivationCell = activationRow.locator('[role=gridcell]').first();
  await plainActivationCell.click();
  const clicked = await node('dg-activation-output').textContent();
  await check('DataGrid pointer activation exact original ID', await json('dg-activation-output'), { rowId: 'a' });
  await plainActivationCell.focus(); await page.keyboard.press('Enter');
  await check('DataGrid keyboard activation exact original ID', await json('dg-activation-output'), { rowId: 'a' });
  const action = rows().first().getByRole('button', { name: 'Review this localized record in detail', exact: true });
  const actionCell = action.locator('xpath=ancestor::*[@role="gridcell"][1]');
  await actionCell.focus(); await page.keyboard.press('Enter');
  await check('DataGrid native group enters interactive cell', await action.evaluate(button => button === document.activeElement), true);
  await page.keyboard.press('Enter');
  await check('DataGrid action exact original identity', await json('dg-action'), { rowId: 'a', actionId: 'review' });
  await check('DataGrid action does not activate another row', await node('dg-activation-output').textContent(), clicked);
  await page.keyboard.press('Escape'); await settle();
  await check('DataGrid native group Escape returns to cell', await actionCell.evaluate(cell => cell === document.activeElement), true);
  const link = rows().first().getByRole('link', { name: 'Alpha', exact: true });
  await link.click();
  await check('DataGrid link emits original identity/value/href', await json('dg-navigation'), { rowId: 'a', value: 'Alpha', href: '#details-a' });
  await check('DataGrid disabled action stays native disabled', await rows().first().getByRole('button', { name: 'Disabled action', exact: true }).isDisabled(), true);
  await activationRow.getByRole('button', { name: 'Custom 0', exact: true }).click();
  await check('DataGrid custom original row state', await activationRow.getByRole('button', { name: 'Custom 1', exact: true }).count(), 1);
  await check('DataGrid custom does not activate row', await node('dg-activation-output').textContent(), clicked);
  const conformance = page.getByRole('button', { name: 'Descriptors, dynamic values and row activation', exact: true });
  await conformance.click();
  await check('DataGrid collapsed conformance panel is inert', await activationRow.evaluate(row => row.closest('.fui-AccordionPanel').hasAttribute('inert')), true);
  await conformance.click();
  await check('DataGrid conformance child state retained after collapse', await activationRow.getByRole('button', { name: 'Custom 1', exact: true }).count(), 1);
  await check('DataGrid person name and avatar vertically centered', await main.locator('.fui-Persona').evaluateAll(elements => elements.every(element => {
    const avatar = element.querySelector('.fui-Avatar').getBoundingClientRect();
    const name = element.querySelector('.fui-Persona__primaryText').getBoundingClientRect();
    return Math.abs((avatar.top + avatar.bottom - name.top - name.bottom) / 2) <= 1;
  })), true);
  const people = rows().first().getByRole('button', { name: 'Bob Example, Carol Example', exact: true });
  await check('DataGrid People native stack overlaps avatar and overflow trigger', await people.evaluate(trigger => {
    const group = trigger.closest('.fui-AvatarGroup');
    const items = group.querySelectorAll(':scope > .fui-AvatarGroupItem');
    if (items.length !== 1) return false;
    const avatar = items[0].getBoundingClientRect(), button = trigger.getBoundingClientRect();
    return Math.min(avatar.right, button.right) - Math.max(avatar.left, button.left) > 0;
  }), true);
  await people.click();
  await check('DataGrid people overflow recovers identities', await page.locator('.fui-PopoverSurface').filter({ hasText: 'Carol Example' }).count() > 0, true);
  await check('DataGrid people overflow one avatar per identity', await page.locator('.fui-AvatarGroupPopover__popoverSurface li').evaluateAll(elements => elements.length === 2 && elements.every(element => element.querySelectorAll('.fui-Avatar').length === 1)), true);
  await check('DataGrid people overflow names vertically centered', await page.locator('.fui-AvatarGroupPopover__popoverSurface li').evaluateAll(elements => elements.every(element => {
    const avatar = element.querySelector('.fui-Avatar').getBoundingClientRect();
    const name = element.querySelector('.fui-Persona__primaryText').getBoundingClientRect();
    return Math.abs((avatar.top + avatar.bottom - name.top - name.bottom) / 2) <= 1;
  })), true);
  await page.locator('.fui-AvatarGroupPopover__popoverSurface').evaluate(element => Promise.all(element.getAnimations({ subtree: true }).map(animation => animation.finished)));
  await page.screenshot({ path: path.join(output, `${label}-data-grid-people-overflow.png`) });
  await page.keyboard.press('Escape');
  await check('DataGrid Tags overflow button follows intrinsic text width', await rows().first().locator('[data-grid-tags] button').evaluate(button => {
    const style = getComputedStyle(button), text = document.createRange();
    text.selectNodeContents(button);
    const intrinsicWidth = text.getBoundingClientRect().width + parseFloat(style.paddingLeft) + parseFloat(style.paddingRight) + parseFloat(style.borderLeftWidth) + parseFloat(style.borderRightWidth);
    return style.minWidth === '0px' && Math.abs(button.getBoundingClientRect().width - intrinsicWidth) <= 1;
  }), true);
  await ratings.first().scrollIntoViewIfNeeded();
  await page.screenshot({ path: path.join(output, `${label}-data-grid-ui-followup.png`) });
  await rows().first().locator('[data-grid-tags] button').click();
  await check('DataGrid configured tag overflow recovers labels', await page.locator('.fui-PopoverSurface').evaluateAll(elements => elements.some(element => element.textContent.includes('Review') && element.textContent.includes('Release'))), true);
  await page.keyboard.press('Escape');
  await rows().first().getByRole('button', { name: /more actions/i }).click();
  await page.getByRole('menuitem', { name: 'Archive record', exact: true }).click();
  await check('DataGrid native menu action original identity', await json('dg-action'), { rowId: 'a', actionId: 'archive' });
  const chart = main.locator('[data-grid-sparkline] svg').first();
  await chart.waitFor();
  await check('DataGrid native chart decoration inert focus', await chart.evaluate(svg => {
    svg.focus(); return document.activeElement !== svg && Boolean(svg.closest('[inert]'));
  }), true);
  const chartCell = chart.locator('xpath=ancestor::*[@role="gridcell"][1]');
  await chartCell.focus();
  await check('DataGrid chart group native focus starts at cell', await chartCell.evaluate(cell => document.activeElement === cell), true);
  const tabFocuses = [];
  for (let step = 0; step < 6; step += 1) {
    await page.keyboard.press('Tab');
    tabFocuses.push(await page.evaluate(() => ({ tag: document.activeElement.tagName,
      graphic: Boolean(document.activeElement.closest('[data-grid-sparkline] svg, [data-grid-sparkline] [inert]')) })));
  }
  await check('DataGrid actual Tab navigation excludes chart graphic nodes', tabFocuses.map(focus => focus.graphic), [false, false, false, false, false, false]);
  await main.locator('[data-grid-sparkline]').first().screenshot({ path: path.join(output, `${label}-data-grid-chart-before.png`) });
  const oldPath = await chart.locator('path[d]').last().getAttribute('d');
  await node('dg-update-sparkline').click(); await settle();
  await check('DataGrid immutable Sparkline points update real chart', await chart.locator('path[d]').last().getAttribute('d') !== oldPath, true);
  await main.locator('[data-grid-sparkline]').first().screenshot({ path: path.join(output, `${label}-data-grid-chart-points-updated.png`) });
  const oldColor = await chart.locator('path[stroke]').last().getAttribute('stroke');
  const colorSvg = await chart.elementHandle();
  const colorPathBefore = await chart.locator('path[stroke]').last().getAttribute('d');
  await node('dg-color').check(); await settle();
  await check('DataGrid Sparkline color updates real chart', await chart.locator('path[stroke]').last().getAttribute('stroke') !== oldColor, true);
  await page.waitForFunction(svg => getComputedStyle(svg.querySelector('path[stroke]')).stroke === 'rgb(209, 52, 56)', colorSvg, { timeout: 2000 });
  await check('DataGrid color-only update retains same native SVG', await chart.evaluate((svg, original) => svg === original, colorSvg), true);
  await check('DataGrid color-only update retains plotted coordinates', await chart.locator('path[stroke]').last().getAttribute('d'), colorPathBefore);
  await check('DataGrid color-only update exact computed paint', await chart.locator('path[stroke]').last().evaluate(path => getComputedStyle(path).stroke), 'rgb(209, 52, 56)');
  await settle();
  await main.locator('[data-grid-sparkline]').first().screenshot({ path: path.join(output, `${label}-data-grid-chart-color-updated.png`) });
  const chartCaptureSession = await page.context().newCDPSession(page);
  const chartSurface = await chartCaptureSession.send('Page.captureScreenshot', { format: 'png', fromSurface: true, captureBeyondViewport: false });
  fs.writeFileSync(path.join(output, `${label}-data-grid-chart-color-updated-cdp.png`), Buffer.from(chartSurface.data, 'base64'));
  await chartCaptureSession.detach();
  await check('DataGrid dynamic missing/invalid fallback', await node('data-grid-dynamic').evaluate(grid => ['Missing text', 'Invalid text', 'Missing number', 'Invalid number'].every(text => grid.textContent.includes(text))), true);
  await check('DataGrid dynamic false/zero survive', await node('data-grid-dynamic').evaluate(grid => grid.textContent.includes('0') && (grid.textContent.toLowerCase().includes('no') || Boolean(grid.querySelector('[aria-label="No"]')))), true);
  await node('dg-resizable').check(); await settle();
  await alignment('resizable');
  const handle = main.locator('.fui-TableResizeHandle').first();
  await handle.scrollIntoViewIfNeeded();
  const box = await handle.boundingBox();
  const beforeWidth = await sortable.evaluate(element => element.getBoundingClientRect().width);
  if (!box) throw new Error('Native Fluent resize handle has no browser geometry.');
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down(); await page.mouse.move(box.x + box.width / 2 + 70, box.y + box.height / 2, { steps: 5 }); await page.mouse.up(); await settle();
  await check('DataGrid real native resize changes column', await sortable.evaluate(element => element.getBoundingClientRect().width) > beforeWidth + 20, true);
  await alignment('after-resize');
  await actionLabelGeometry('after-resize');
  const responsiveGrid = node('dg-responsive-tags-grid');
  const responsiveTags = responsiveGrid.locator('[data-grid-tags]');
  const visibleTags = () => responsiveTags.locator('.fui-Tag').evaluateAll(tags => tags.filter(tag => getComputedStyle(tag).display !== 'none' && getComputedStyle(tag).visibility !== 'hidden' && tag.getClientRects().length > 0).map(tag => tag.textContent));
  await check('DataGrid wide responsive fixture exposes all tags', await visibleTags(), ['Design', 'Review', 'Release']);
  await node('dg-responsive-tags-width').selectOption('100');
  await settle();
  await page.waitForFunction(() => Boolean(document.querySelector('#dg-responsive-tags-grid [data-overflowing]')));
  await check('DataGrid native ResizeObserver narrows responsive Tags column', await responsiveTags.evaluate(tags => {
    const cell = tags.closest('[role=gridcell]').getBoundingClientRect();
    return cell.width <= 101;
  }), true);
  const hiddenLabels = await responsiveTags.locator('[data-overflowing]').evaluateAll(tags => tags.map(tag => tag.textContent));
  await check('DataGrid responsive Tags have genuinely hidden items', hiddenLabels.length > 0, true);
  await check('DataGrid responsive Tags use native hidden display geometry', await responsiveTags.locator('[data-overflowing]').evaluateAll(tags => tags.every(tag => getComputedStyle(tag).display === 'none' && tag.getClientRects().length === 0)), true);
  await responsiveTags.getByRole('button').click();
  await check('DataGrid responsive Tags popup recovers every hidden label', await page.locator('.fui-PopoverSurface').evaluateAll((popovers, labels) => popovers.some(popover => labels.every(label => popover.textContent.includes(label))), hiddenLabels), true);
  await page.keyboard.press('Escape'); await settle();
  await node('dg-responsive-tags-width').selectOption('300'); await settle();
  await page.waitForFunction(() => document.querySelectorAll('#dg-responsive-tags-grid [data-overflowing]').length === 0);
  await check('DataGrid restored available width restores all responsive Tags', await visibleTags(), ['Design', 'Review', 'Release']);
  await capture('light');
  for (const theme of ['dark', 'highContrast']) {
    await page.evaluate(theme => window.demoFixtureHost.update({ theme: theme === 'highContrast' ? 'contrast' : theme }), theme); await node('dg-dir').selectOption('rtl'); await settle();
    await check(`DataGrid ${theme} RTL direction`, await main.evaluate(element => getComputedStyle(element).direction), 'rtl');
    await alignment(`${theme}-rtl`); await capture(`${theme}-rtl`);
  }
  await node('dg-width').selectOption('320'); await page.setViewportSize({ width: 420, height: 900 }); await settle();
  await check('DataGrid narrow local scroll region', await main.evaluate(grid => {
    let element = grid.parentElement;
    while (element && element !== document.body) {
      if (['auto', 'scroll'].includes(getComputedStyle(element).overflowX) && element.scrollWidth > element.clientWidth) return element.clientWidth <= 321;
      element = element.parentElement;
    }
    return false;
  }), true);
  await check('DataGrid narrow page has no horizontal overflow', await page.evaluate(() => document.documentElement.scrollWidth <= document.documentElement.clientWidth + 1), true);
  await actionLabelGeometry('narrow-scroll');
  await capture('narrow-rtl');
  await page.emulateMedia({ forcedColors: 'active' });
  await firstCell.focus();
  await check('DataGrid forced colors native focus survives', await firstCell.evaluate(element => element === document.activeElement && getComputedStyle(element).visibility === 'visible'), true);
  await capture('forced-colors'); await page.emulateMedia({ forcedColors: 'none' });
  // Pinch scaling and desktop-layout emulation are distinct evidence. Neither
  // operation proves Chrome's genuine desktop browser UI zoom setting.
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Emulation.setPageScaleFactor', { pageScaleFactor: 2 });
  await check('DataGrid pinch zoom 200% visual viewport scale', await page.evaluate(() => window.visualViewport.scale), 2);
  await capture('pinch200');
  await cdp.send('Emulation.setPageScaleFactor', { pageScaleFactor: 1 });
  await node('dg-width').selectOption('1100');
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 720, height: 500, deviceScaleFactor: 2, mobile: false });
  await settle();
  await check('DataGrid 200% desktop layout equivalent measured CSS viewport and DPR', await page.evaluate(() => ({ width: innerWidth, height: innerHeight, dpr: devicePixelRatio, scale: visualViewport.scale })), { width: 720, height: 500, dpr: 2, scale: 1 });
  await check('DataGrid 200% layout equivalent page reflow stays bounded', await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true);
  await check('DataGrid 200% layout equivalent local scroll stays usable', await main.evaluate(grid => {
    let region = grid.parentElement;
    while (region && region !== document.body) {
      if (['auto', 'scroll'].includes(getComputedStyle(region).overflowX) && region.scrollWidth > region.clientWidth) return region.clientWidth <= innerWidth && region.clientWidth > 0;
      region = region.parentElement;
    }
    return false;
  }), true);
  await firstCell.focus(); await page.keyboard.press('ArrowDown');
  await check('DataGrid 200% layout equivalent keyboard navigation usable', await rows().nth(1).evaluate(row => row.contains(document.activeElement)), true);
  await alignment('layout-equivalent200'); await actionLabelGeometry('layout-equivalent200');
  await capture('desktop-layout-equivalent200');
  await cdp.send('Emulation.clearDeviceMetricsOverride'); await cdp.detach();
  await page.setViewportSize({ width: 1440, height: 1000 }); await node('dg-width').selectOption('1100');
  const mountedBody = await main.locator('.fui-DataGridBody').elementHandle();
  const mountedRow = await rows().first().elementHandle();
  const mainCustom = rows().first().locator('[data-custom-row=a]');
  await mainCustom.click();
  const mountedCustom = await mainCustom.elementHandle();
  const selectedBeforeRefresh = await selectedIds();
  await check('DataGrid body has one always-mounted AutoSkeleton content wrapper', await main.locator('[data-auto-skeleton-content]').count(), 1);
  await node('dg-loading').check(); await settle();
  if (await main.locator('[data-auto-skeleton-content]').count() === 1) {
    await page.waitForFunction(() => Boolean(document.querySelector('#data-grid-main [data-auto-skeleton-overlay] [data-auto-skeleton-shape]')));
  }
  await check('DataGrid refresh retains supplied rows', await rows().evaluateAll(elements => elements.filter(row => !row.hasAttribute('data-grid-loading-row')).map(row => row.querySelector('[data-grid-column-id=title]')?.textContent)), initialNames);
  await check('DataGrid refresh pads two supplied rows to minimum four', await rows().count(), 4);
  await check('DataGrid refresh native body DOM identity retained', await main.locator('.fui-DataGridBody').evaluate((body, previous) => body === previous, mountedBody), true);
  await check('DataGrid refresh supplied row DOM identity retained', await rows().first().evaluate((row, previous) => row === previous, mountedRow), true);
  await check('DataGrid refresh consumer child DOM identity retained', await mainCustom.evaluate((button, previous) => button === previous, mountedCustom), true);
  await check('DataGrid refresh consumer child state retained', await mainCustom.textContent(), 'Custom 1');
  await check('DataGrid loading native busy semantics', await main.getAttribute('aria-busy'), 'true');
  await check('DataGrid refresh AutoSkeleton content hidden and inert', await main.locator('[data-auto-skeleton-content]').evaluateAll(elements => elements.length === 1 && elements.every(content => content.hasAttribute('inert') && content.getAttribute('aria-hidden') === 'true' && getComputedStyle(content).opacity === '0')), true);
  await check('DataGrid refresh has one body AutoSkeleton overlay', await main.locator('[data-auto-skeleton-overlay]').count(), 1);
  await check('DataGrid refresh suppresses all native selection cells', await main.locator('.fui-DataGridSelectionCell').count(), 0);
  await check('DataGrid refresh does not emit synthetic selection IDs', await selectedIds(), selectedBeforeRefresh);
  await check('DataGrid refresh has no outer loading progress indicator', await node('data-grid-demo').getByRole('progressbar', { name: 'Loading', exact: true }).count(), 0);
  const refreshGeometry = await loadingCellSkeletons('refresh');
  const refreshFingerprint = loadingFingerprint(refreshGeometry);
  const loadingTheme = await page.evaluate(() => window.demoFixtureHost.config.theme);
  await node('dg-width').selectOption('640'); await settle();
  await check('DataGrid simulated loading values remain stable during width rerender', loadingFingerprint(await loadingGeometry()), refreshFingerprint);
  await page.evaluate(() => window.demoFixtureHost.update({ theme: 'light' })); await settle();
  await check('DataGrid simulated loading values remain stable during theme rerender', loadingFingerprint(await loadingGeometry()), refreshFingerprint);
  await check('DataGrid simulated loading bones remain bounded after responsive rerender', loadingBoundsMatch(await loadingGeometry()), true);
  await capture('refresh-loading');
  await node('dg-width').selectOption('1100'); await page.evaluate(theme => window.demoFixtureHost.update({ theme }), loadingTheme); await settle();
  await node('dg-loading').uncheck(); await settle();
  await check('DataGrid refresh completion removes padding rows', await rows().count(), 2);
  await check('DataGrid refresh completion retains body DOM identity', await main.locator('.fui-DataGridBody').evaluate((body, previous) => body === previous, mountedBody), true);
  await check('DataGrid refresh completion restores native selection cells', await main.locator('.fui-DataGridSelectionCell').count(), 3);
  await check('DataGrid refresh completion preserves controlled selected IDs', await selectedIds(), selectedBeforeRefresh);
  await check('DataGrid refresh completion clears overlay and inert content', await main.evaluate(grid => grid.querySelectorAll('[data-auto-skeleton-overlay]').length === 0 && grid.querySelectorAll('[data-auto-skeleton-content][inert]').length === 0), true);
  await check('DataGrid refresh completion removes simulated values and restores native cell markup', await main.locator('.fui-DataGridBody [data-grid-column-id]').evaluateAll(cells => cells.length === 32 && cells.every(cell => {
    const style = getComputedStyle(cell);
    return !cell.querySelector('[data-grid-loading-kind]') && !cell.hasAttribute('data-skeleton-shape') && [style.borderTopLeftRadius, style.borderTopRightRadius, style.borderBottomLeftRadius, style.borderBottomRightRadius].every(radius => radius === '0px');
  })), true);
  await check('DataGrid refresh completion retains custom state and DOM identity', await mainCustom.evaluate((button, previous) => button === previous && button.textContent === 'Custom 1', mountedCustom), true);
  await node('dg-loading').check(); await settle();
  if (await main.locator('[data-auto-skeleton-content]').count() === 1) {
    await page.waitForFunction(() => Boolean(document.querySelector('#data-grid-main [data-auto-skeleton-overlay] [data-auto-skeleton-shape]')));
  }
  await check('DataGrid simulated loading values remain stable across loading toggles', loadingFingerprint(await loadingGeometry()), refreshFingerprint);
  await node('dg-loading').uncheck(); await settle();
  await node('dg-error').check();
  await check('DataGrid error retains rows', await rowNames(), initialNames);
  await node('dg-inspect-slot-refs').click();
  await check('DataGrid error inner slot ref stays native div', await node('dg-slot-refs').textContent(), 'header=SPAN error=DIV missing=SPAN');

  const retries = Number(await node('dg-retries').textContent());
  await node('data-grid-demo').getByRole('button', { name: 'Retry', exact: true }).click();
  await check('DataGrid retry is consumer intent', Number(await node('dg-retries').textContent()), retries + 1);
  await node('dg-empty').check();
  await check('DataGrid empty replaces native grid', await main.count(), 0);
  const emptyStatus = node('data-grid-demo').getByRole('status').filter({ hasText: /^No data available$/ });
  await check('DataGrid default empty text is horizontally centered', await emptyStatus.evaluate(status => {
    const range = document.createRange(); range.selectNodeContents(status);
    const text = range.getBoundingClientRect(), container = status.getBoundingClientRect();
    return Math.abs((text.left + text.right - container.left - container.right) / 2) <= 1;
  }), true);
  await check('DataGrid default empty uses native compact block padding', await emptyStatus.evaluate(status => {
    const style = getComputedStyle(status), token = style.getPropertyValue('--spacingVerticalM').trim();
    return token === '12px' && style.paddingTop === token && style.paddingBottom === token;
  }), true);
  await capture('default-empty');
  await node('dg-inspect-ref').click(); await check('DataGrid empty native ref absent', await node('dg-ref').textContent(), 'absent');
  await node('dg-empty-state').selectOption('noResults');
  await check('DataGrid no-results status', await node('data-grid-demo').getByRole('status').filter({ hasText: /no results/i }).count() > 0, true);
  const noResultsStatus = node('data-grid-demo').getByRole('status').filter({ hasText: /^No results found$/ });
  await check('DataGrid default no-results text is horizontally centered', await noResultsStatus.evaluate(status => {
    const range = document.createRange(); range.selectNodeContents(status);
    const text = range.getBoundingClientRect(), container = status.getBoundingClientRect();
    return Math.abs((text.left + text.right - container.left - container.right) / 2) <= 1;
  }), true);
  await check('DataGrid default no-results uses native compact block padding', await noResultsStatus.evaluate(status => {
    const style = getComputedStyle(status), token = style.getPropertyValue('--spacingVerticalM').trim();
    return token === '12px' && style.paddingTop === token && style.paddingBottom === token;
  }), true);
  await capture('default-no-results');
  await node('dg-custom-status-slots').check();
  await check('DataGrid custom no-results callback visible', await node('data-grid-demo').getByRole('status').filter({ hasText: /^Consumer-owned no-results slot$/ }).count(), 1);
  await node('dg-slot-suppress').check();
  await check('DataGrid explicit false suppresses custom status', await node('dg-frame').locator('[role=status]').textContent(), '');
  await node('dg-slot-suppress').uncheck();
  await node('dg-empty-state').selectOption('empty');
  await check('DataGrid custom empty slot visible', await node('data-grid-demo').getByRole('status').filter({ hasText: /^Consumer-owned empty slot$/ }).count(), 1);
  await node('dg-custom-status-slots').uncheck();
  await node('dg-empty-state').selectOption('noResults');

  await node('dg-loading').check(); await settle();
  if (await main.locator('[data-auto-skeleton-content]').count() === 1) {
    await page.waitForFunction(() => Boolean(document.querySelector('#data-grid-main [data-auto-skeleton-overlay] [data-auto-skeleton-shape]')));
  }
  await check('DataGrid initial loading minimum four rows', await rows().count(), 4);
  await check('DataGrid initial loading contains four placeholder rows', await main.locator('.fui-DataGridRow[data-grid-loading-row]').count(), 4);
  await check('DataGrid initial loading suppresses all native selection cells', await main.locator('.fui-DataGridSelectionCell').count(), 0);
  await check('DataGrid initial loading has one measured AutoSkeleton overlay', await main.locator('[data-auto-skeleton-overlay]').count(), 1);
  await check('DataGrid initial loading content hidden and inert', await main.locator('[data-auto-skeleton-content]').evaluateAll(elements => elements.length === 1 && elements.every(content => content.hasAttribute('inert') && getComputedStyle(content).opacity === '0')), true);
  await check('DataGrid initial loading has no visible Loading status', await node('data-grid-demo').getByRole('status').filter({ hasText: /^Loading$/ }).count(), 0);
  await loadingCellSkeletons('initial loading');
  await capture('initial-loading');
  await node('dg-loading').uncheck(); await node('dg-empty').uncheck(); await node('dg-hidden').check();
  await check('DataGrid all-hidden replaces native grid', await main.count(), 0);
  await node('dg-inspect-ref').click(); await check('DataGrid all-hidden native ref absent', await node('dg-ref').textContent(), 'absent');
  await node('dg-hidden').uncheck(); await node('dg-inspect-ref').click();
  await check('DataGrid restored ref native element', await node('dg-ref').textContent(), 'div');
  const loadingBoundary = await node('dg-loading-boundaries-grid').evaluateAll(grids => {
    const grid = grids[0];
    if (grids.length !== 1) return { nativeColumns: false, boundedTemplates: false, measuredFallback: false };
    const cells = [...grid.querySelectorAll('.fui-DataGridBody [data-grid-column-id]')];
    const shapes = cells.flatMap(cell => [...cell.querySelectorAll('[data-grid-loading-kind] [data-skeleton-shape=rectangle]')]);
    const bones = [...grid.querySelectorAll('[data-auto-skeleton-overlay] [data-auto-skeleton-shape]')];
    const eligible = shapes.filter(shape => { const rect = shape.getBoundingClientRect(); return Math.round(rect.width) >= 8 && Math.round(rect.height) >= 8; });
    const wideShapes = cells.filter(cell => cell.getAttribute('data-grid-column-id') === 'title').flatMap(cell => [...cell.querySelectorAll('[data-skeleton-shape=rectangle]')]);
    return {
      nativeColumns: cells.length === 20 && cells.every(cell => {
        const rect = cell.getBoundingClientRect(), style = getComputedStyle(cell), contentWidth = cell.getAttribute('data-grid-column-id') === 'title' ? 150 : 20;
        const outerWidth = contentWidth + parseFloat(style.paddingLeft) + parseFloat(style.paddingRight) + parseFloat(style.borderLeftWidth) + parseFloat(style.borderRightWidth);
        return [rect.left, rect.top, rect.width, rect.height].every(Number.isFinite) && parseFloat(style.width) === contentWidth && Math.abs(rect.width - outerWidth) <= 1;
      }),
      boundedTemplates: cells.length === 20 && cells.every(cell => {
        const rect = cell.getBoundingClientRect(), style = getComputedStyle(cell);
        const primitives = [...cell.querySelectorAll('[data-grid-loading-kind] [data-skeleton-shape=rectangle]')].map(shape => shape.getBoundingClientRect());
        return cell.querySelectorAll('[data-grid-loading-kind]').length === 1 && primitives.length > 0 && primitives.every((shape, index) =>
          [shape.left, shape.top, shape.width, shape.height].every(Number.isFinite) && shape.width > 0 && shape.height > 0 &&
          shape.left >= rect.left + parseFloat(style.paddingLeft) - 0.1 && shape.right <= rect.right - parseFloat(style.paddingRight) + 0.1 && shape.top >= rect.top - 0.1 && shape.bottom <= rect.bottom + 0.1 &&
          primitives.slice(index + 1).every(other => Math.min(shape.right, other.right) - Math.max(shape.left, other.left) <= 0.1 || Math.min(shape.bottom, other.bottom) - Math.max(shape.top, other.top) <= 0.1));
      }),
      // Existing AutoSkeleton deliberately ignores normalized leaves below 8px.
      measuredFallback: wideShapes.length === 4 && wideShapes.every(shape => eligible.includes(shape)) && shapes.length > eligible.length && bones.length === eligible.length && grid.querySelectorAll('[data-auto-skeleton-overlay]').length === 1 &&
        bones.every(bone => bone.getAttribute('data-auto-skeleton-shape') === 'rectangle') && eligible.every(shape => {
          const rect = shape.getBoundingClientRect();
          const radius = Math.min(parseFloat(getComputedStyle(shape).getPropertyValue('--borderRadiusMedium')), Math.round(rect.width) / 2, Math.round(rect.height) / 2);
          return bones.some(bone => {
            const box = bone.getBoundingClientRect(), style = getComputedStyle(bone);
            return ['left', 'top', 'width', 'height'].every(property => Math.abs(box[property] - rect[property]) <= 1) &&
              [style.borderTopLeftRadius, style.borderTopRightRadius, style.borderBottomLeftRadius, style.borderBottomRightRadius].every(corner => parseFloat(corner) === radius);
          });
        }),
      nativeWidths: [...new Set(cells.map(cell => cell.getBoundingClientRect().width))],
      primitiveCount: shapes.length, boneCount: bones.length, eligibleCount: eligible.length,
      subminimumWidths: [...new Set(shapes.map(shape => shape.getBoundingClientRect().width).filter(value => Math.round(value) < 8))],
    };
  });
  await check('DataGrid narrow loading fixture preserves finite native 20px content columns', loadingBoundary.nativeColumns, true);
  await check('DataGrid narrow compound loading templates stay bounded without overlap', loadingBoundary.boundedTemplates, true);
  await check('DataGrid narrow loading honors existing AutoSkeleton minimum-size fallback', loadingBoundary.measuredFallback, true);
  fs.writeFileSync(path.join(output, `${label}-data-grid-loading-boundaries-geometry.json`), JSON.stringify(loadingBoundary, null, 2));
  const unmeasurable = node('dg-loading-unmeasurable-grid');
  await check('DataGrid unmeasurable templates preserve busy semantics and four bounded tiny cells', await unmeasurable.evaluateAll(grids => grids.length === 1 && grids.every(grid => {
    const cells = [...grid.querySelectorAll('.fui-DataGridBody [data-grid-column-id]')];
    return grid.getAttribute('aria-busy') === 'true' && grid.querySelectorAll('.fui-DataGridBody > [role=row]').length === 4 && cells.length === 4 && cells.every(cell => {
      const box = cell.getBoundingClientRect(), style = getComputedStyle(cell);
      const shapes = [...cell.querySelectorAll('[data-grid-loading-kind=person] [data-skeleton-shape=rectangle]')].map(shape => shape.getBoundingClientRect());
      return Math.abs(box.width - 20) <= 1 && shapes.length === 2 && shapes.every((shape, index) => [shape.left, shape.top, shape.width, shape.height].every(Number.isFinite) &&
        shape.width > 0 && Math.round(shape.width) < 8 && shape.left >= box.left + parseFloat(style.paddingLeft) - 0.1 && shape.right <= box.right - parseFloat(style.paddingRight) + 0.1 && shape.top >= box.top - 0.1 && shape.bottom <= box.bottom + 0.1 &&
        shapes.slice(index + 1).every(other => Math.min(shape.right, other.right) - Math.max(shape.left, other.left) <= 0.1 || Math.min(shape.bottom, other.bottom) - Math.max(shape.top, other.top) <= 0.1));
    });
  })), true);
  await check('DataGrid unmeasurable templates keep actual fallback content available', await unmeasurable.evaluateAll(grids => grids.length === 1 && grids.every(grid => {
    const content = grid.querySelector('[data-auto-skeleton-content]'), actual = grid.querySelector('[data-grid-row-id=actual] .fui-Persona__primaryText');
    const live = actual?.closest('[data-skeleton-ignore]');
    return grid.querySelectorAll('[data-auto-skeleton-overlay]').length === 0 && content && !content.hasAttribute('inert') && content.getAttribute('aria-hidden') !== 'true' && getComputedStyle(content).opacity === '1' && actual?.textContent === 'Visible fallback owner' &&
      getComputedStyle(actual).display !== 'none' && actual.getBoundingClientRect().height > 0 && live && getComputedStyle(live).display === 'contents';
  })), true);
  await check('DataGrid unmeasurable fallback exposes the original accessible name', await unmeasurable.getByText('Visible fallback owner', { exact: true }).isVisible() && await unmeasurable.getByRole('img', { name: 'Visible fallback owner', exact: true }).count() === 1, true);
  await node('dg-loading-boundaries').scrollIntoViewIfNeeded();
  await page.screenshot({ path: path.join(output, `${label}-data-grid-loading-boundaries.png`) });
  await exerciseNativeDataGridZoom({ url, output, label, boundary, check, consoleMessages });
  await check('DataGrid console health', consoleMessages, []);
}
/** Verify Chrome native zoom using only disposable profiles and measured metrics. */
async function exerciseNativeDataGridZoom({ url, output, label, boundary, check, consoleMessages }) {
  const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
  const executablePath = process.env.SX_BROWSER_EXECUTABLE || process.env.CHROME_EXECUTABLE_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
  const metrics = [];
  for (const zoom of [1, 2]) {
    const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'data-grid-native-zoom-'));
    fs.mkdirSync(path.join(profile, 'Default'));
    // Chromium's kPartitionDefaultZoomLevel is a dictionary. The default
    // partition key is "x", as documented by chrome_zoom_level_prefs.cc.
    fs.writeFileSync(path.join(profile, 'Default', 'Preferences'), JSON.stringify({ partition: { default_zoom_level: { x: Math.log(zoom) / Math.log(1.2) } } }));
    let context;
    const pageErrors = [];
    try {
      context = await chromium.launchPersistentContext(profile, { executablePath, headless: true, viewport: null, args: ['--window-size=1440,1000'] });
      const page = context.pages()[0];
      page.on('console', message => { if (['error', 'warning'].includes(message.type())) consoleMessages.push({ type: message.type(), text: message.text() }); });
      page.on('pageerror', error => pageErrors.push(error.message));
      await page.goto(`${url}/?data-grid-only=1`);
      const main = page.locator('#data-grid-main');
      await main.waitFor();
      metrics.push(await page.evaluate(() => ({ width: innerWidth, height: innerHeight, dpr: devicePixelRatio, scale: visualViewport.scale })));
      await check(`DataGrid native zoom ${zoom * 100}% boundary`, await page.locator('[data-grid-boundary]').getAttribute('data-grid-boundary'), boundary);
      if (zoom === 2) {
        await check('DataGrid measured native zoom doubles DPR and halves CSS viewport', {
          dprRatio: metrics[1].dpr / metrics[0].dpr,
          widthRatio: metrics[0].width / metrics[1].width,
          heightRounding: Math.abs(metrics[0].height - metrics[1].height * 2) <= 1,
          visualScale: metrics[1].scale,
        }, { dprRatio: 2, widthRatio: 2, heightRounding: true, visualScale: 1 });
        await page.evaluate(() => window.demoFixtureHost.update({ theme: 'dark' })); await page.locator('#dg-dir').selectOption('rtl');
        await check('DataGrid native zoom200 controls reflow stays bounded', await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true);
        await check('DataGrid native zoom200 scroll region fits CSS viewport', await main.evaluate(grid => {
          const region = grid.closest('[data-grid-scroll]');
          return region.clientWidth <= innerWidth && region.clientWidth > 0 && region.scrollWidth > region.clientWidth;
        }), true);
        const firstCell = main.locator('[data-grid-row-id="a"] [data-grid-column-id="title"]');
        await firstCell.focus(); await page.keyboard.press('ArrowDown');
        await check('DataGrid native zoom200 keyboard navigation usable', await main.locator('[data-grid-row-id="b"]').evaluate(row => row.contains(document.activeElement)), true);
        const action = main.locator('[data-grid-row-id="a"]').getByRole('button', { name: 'Review this localized record in detail', exact: true });
        await action.click();
        await check('DataGrid native zoom200 action original row identity', JSON.parse(await page.locator('#dg-action').textContent()), { rowId: 'a', actionId: 'review' });
        await check('DataGrid native zoom200 long label stays inside Actions cell', await action.evaluate(button => {
          const cell = button.closest('[role=gridcell]').getBoundingClientRect();
          const range = document.createRange(); range.selectNodeContents(button);
          return [...range.getClientRects()].every(rect => rect.left >= cell.left - 1 && rect.right <= cell.right + 1 && rect.top >= cell.top - 1 && rect.bottom <= cell.bottom + 1);
        }), true);
        await check('DataGrid native zoom200 header/body remain aligned', await main.evaluate(grid => {
          const head = [...grid.querySelectorAll('.fui-DataGridHeader [data-grid-column-id]')];
          const body = [...grid.querySelectorAll('[data-grid-row-id="a"] > [data-grid-column-id]')];
          return head.length === 16 && body.length === 16 && head.every((element, index) => {
            const h = element.getBoundingClientRect(), b = body[index].getBoundingClientRect();
            return Math.abs(h.width - b.width) <= 1 && Math.abs(h.left - b.left) <= 1;
          });
        }), true);
        await page.screenshot({ path: path.join(output, `${label}-data-grid-native-zoom200-dark-rtl.png`), fullPage: true });
        await page.evaluate(() => window.scrollTo(0, 0));
        await page.screenshot({ path: path.join(output, `${label}-data-grid-native-zoom200-dark-rtl-viewport.png`), fullPage: false });
        const captureSession = await context.newCDPSession(page);
        const nativeCapture = await captureSession.send('Page.captureScreenshot', { format: 'png', fromSurface: true, captureBeyondViewport: false });
        fs.writeFileSync(path.join(output, `${label}-data-grid-native-zoom200-cdp.png`), Buffer.from(nativeCapture.data, 'base64'));
        await captureSession.detach();
        fs.writeFileSync(path.join(output, `${label}-data-grid-native-zoom200.html`), await page.content());
      }
      await check(`DataGrid native zoom ${zoom * 100}% page errors`, pageErrors, []);
    } finally {
      if (context) await context.close();
      fs.rmSync(profile, { recursive: true, force: true });
    }
  }
  fs.writeFileSync(path.join(output, `${label}-native-zoom-metrics.json`), JSON.stringify({ windowSize: [1440, 1000], viewport: 'native window', preference: 'partition.default_zoom_level.x', metrics }, null, 2));
}
/** Separate fixture-only rendering fault; this is not a native valid-input Charts failure. */
async function exerciseDataGridChartProbe({ page, url, output, label, boundary, check, consoleMessages, pageErrors }) {
  const message = 'Injected DataGrid chart failure (fixture-only)';
  const main = page.locator('#data-grid-main');
  const firstCell = main.locator('[data-grid-row-id="a"] [data-grid-column-id="trend"]');
  const charts = main.locator('[data-grid-sparkline] svg');
  const settle = () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  const capture = async phase => { await firstCell.screenshot({ path: path.join(output, `${label}-chart-fault-${phase}.png`) }); };
  await page.goto(`${url}/?data-grid-only=1`);
  await main.waitFor();
  await check('Chart fault probe explicit library boundary', await page.locator('[data-grid-boundary]').getAttribute('data-grid-boundary'), boundary);
  await check('Chart fault probe starts with two actual Charts SVGs', await charts.count(), 2);
  await page.evaluate(() => window.dataGridChartProbe.setFailure(true));
  await page.locator('#dg-color').check(); await settle();
  const retry = firstCell.getByRole('button', { name: 'Retry', exact: true });
  await retry.waitFor();
  await check('Chart fault probe isolates error without losing grid rows', await main.locator('[data-grid-row-id]').count(), 2);
  await check('Chart fault probe summary survives failed rendering', await firstCell.getByRole('img').evaluate(image => Boolean(image.getAttribute('aria-label')?.includes('last 7'))), true);
  await check('Chart fault probe replaces failed SVG', await firstCell.locator('svg').count(), 0);
  await capture('fallback');
  await page.evaluate(() => window.dataGridChartProbe.setFailure(false));
  await firstCell.focus(); await page.keyboard.press('Enter'); await settle();
  await check('Chart fault probe native group enters retry button', await retry.evaluate(button => document.activeElement === button), true);
  await page.keyboard.press('Enter'); await settle();
  await check('Chart fault probe keyboard retry restores actual Charts SVG', await firstCell.locator('svg').count(), 1);
  await check('Chart fault probe keyboard retry clears fallback', await firstCell.getByRole('button', { name: 'Retry', exact: true }).count(), 0);
  await capture('keyboard-retry');
  await page.evaluate(() => window.dataGridChartProbe.setFailure(true));
  await page.locator('#dg-color').uncheck(); await settle();
  await firstCell.getByRole('button', { name: 'Retry', exact: true }).waitFor();
  await page.evaluate(() => window.dataGridChartProbe.setFailure(false));
  await page.locator('#dg-update-sparkline').click(); await settle();
  await check('Chart fault probe changed observations reset both boundaries', await charts.count(), 2);
  await check('Chart fault probe reset keeps updated accessible summary', await firstCell.getByRole('img').first().evaluate(image => Boolean(image.getAttribute('aria-label')?.includes('last 17'))), true);
  await check('Chart fault probe reset clears retry controls', await main.getByRole('button', { name: 'Retry', exact: true }).count(), 0);
  await capture('observations-reset');
  const injected = await page.evaluate(() => window.dataGridChartProbe.injectedFailures);
  await check('Chart fault probe confirms fixture injection occurred', injected > 0, true);
  const expectedReactDiagnostic = entry => entry.type === 'error' && entry.text.startsWith('The above error occurred in the <Sparkline> component:') && entry.text.includes('ChartBoundary') && entry.text.includes('DataGridSparklineCell');
  await check('Chart fault probe expected handled React diagnostics occurred', consoleMessages.filter(expectedReactDiagnostic).length > 0, true);
  await check('Chart fault probe rejects unrelated console diagnostics', consoleMessages.filter(entry => !expectedReactDiagnostic(entry)), []);
  await check('Chart fault probe expected injected browser errors occurred', pageErrors.filter(error => error === message).length > 0, true);
  await check('Chart fault probe rejects unrelated page errors', pageErrors.filter(error => error !== message), []);
  fs.writeFileSync(path.join(output, `${label}-chart-fault-provenance.json`), JSON.stringify({ faultInjection: 'Fixture-only wrapper throws before official Charts rendering; successful renders delegate to actual Charts', nativeValidInputFailure: 'NOT EXECUTED', boundary, injected, handledReactDiagnostics: consoleMessages.length, injectedPageErrors: pageErrors.length }, null, 2));
}
module.exports = { exerciseDataGrid, exerciseDataGridChartProbe };
