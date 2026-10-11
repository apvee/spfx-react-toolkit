// Both routes exercise the actual lazy-sample panel with real Fluent and Form.
// Browser plugin not available; reuse installed Playwright, no dependency install.
const fs = require('node:fs');
const path = require('node:path');

async function checkStageBrowser(page, { url, output, label, boundary, check, consoleMessages }) {
  const instance = (id = 'A') => page.locator(`[data-stage-demo="instance"][data-stage-instance="${id}"]`);
  const surface = (id = 'A') => page.locator(`[data-stage-demo="surface"][data-stage-instance="${id}"]:visible`);
  const marker = (root, name) => root.locator(`[data-stage-demo="${name}"]`);
  const settle = async () => {
    await page.evaluate(async () => {
      const paint = () => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      await paint();
      const documents = [document, ...[...document.querySelectorAll('[data-stage-probe="iframe"]')].map(frame => frame.contentDocument).filter(Boolean)];
      await Promise.all(documents.flatMap(owner => owner.getAnimations()).filter(animation => animation.effect?.getTiming().iterations !== Infinity).map(animation => animation.finished.catch(() => undefined)));
      await paint();
    });
  };
  const capture = async name => {
    await settle();
    await page.screenshot({ path: path.join(output, `${label}-stage-${name}.png`), fullPage: false });
    fs.writeFileSync(path.join(output, `${label}-stage-${name}.html`), await page.content());
  };
  const cases = [];
  const observations = [];
  const scenario = async (name, run) => {
    await page.setViewportSize({ width: 1210, height: 900 });
    await page.emulateMedia({ forcedColors: 'none' });
    await page.goto(url);
    await marker(instance(), 'trigger').waitFor();
    try {
      await run();
      cases.push({ name, completed: true });
    } catch (error) {
      cases.push({ name, completed: false, error: error.message });
      await check(`stage:${name}:completed`, error.message, 'completed');
      fs.writeFileSync(path.join(output, `${label}-stage-${name}-failure.txt`), error.stack);
      await capture(`${name}-failure`);
    }
  };
  const configure = async (mode, size = 'small', options = {}, id = 'A') => {
    const owner = instance(id);
    await marker(owner, 'mode').selectOption(mode);
    await marker(owner, 'size').selectOption(size);
    for (const [name, value] of Object.entries(options)) {
      if (typeof value === 'boolean') await marker(owner, name).setChecked(value);
      else await marker(owner, name).selectOption(value);
    }
  };
  const open = async (id = 'A') => {
    await marker(instance(id), 'trigger').click();
    await surface(id).waitFor(); await settle();
  };
  const close = async (id = 'A') => {
    await surface(id).getByRole('button', { name: `Close ${id}`, exact: true }).click();
    await marker(instance(id), 'open').filter({ hasText: 'Open: false' }).waitFor();
    await settle();
  };
  const rectangle = async target => target.evaluate(element => {
    const rect = element.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height, right: rect.right, bottom: rect.bottom };
  });
  const approximately = (value, expected) => Math.abs(value - expected) <= 1;

  await scenario('identity', async () => {
    await check('stage:page-title', await page.title(), `Stage ${boundary} browser checkpoint`);
    await check('stage:page-url', new URL(page.url()).origin, url);
    await check('stage:real-sample-panel', await page.locator('[data-stage-demo="panel"]').count(), 1);
    await check('stage:boundary', await page.locator('[data-stage-fixture="boundary"]').textContent(), boundary);
    await check('stage:framework-overlay', await page.locator('webpack-dev-server-client-overlay, #webpack-dev-server-client-overlay, nextjs-portal').count(), 0);
  });

  for (const [mode, widths] of Object.entries({ dialog: [640, 800, 1120], drawer: [320, 592, 940], popover: [240, 320, 480] })) {
    await scenario(`${mode}-geometry`, async () => {
      for (const [index, size] of ['small', 'medium', 'large'].entries()) {
        await configure(mode, size); await open();
        for (const viewport of [1920, 1000, 390]) {
          await page.setViewportSize({ width: viewport, height: 900 }); await settle();
          const rect = await rectangle(surface());
          // Native Fluent scroll lock reserves the classic scrollbar gutter. An
          // independent 100vw/100vh class measures usable CSS viewport units,
          // rather than inferring them from Stage or physical window.innerWidth.
          const cssViewport = await rectangle(page.locator('[data-stage-probe="viewport"]'));
          const layoutViewport = await page.evaluate(() => ({ width: document.documentElement.clientWidth, height: document.documentElement.clientHeight }));
          const gap = Math.min(16, Math.max(8, cssViewport.width / 100));
          const availableWidth = mode === 'popover' ? Math.min(cssViewport.width, layoutViewport.width) : cssViewport.width;
          const expectedWidth = Math.min(widths[index], availableWidth - (mode === 'dialog' ? gap * 2 : 0));
          await check(`stage:${mode}:${size}:${viewport}:width`, approximately(rect.width, expectedWidth), true);
          await check(`stage:${mode}:${size}:${viewport}:no-horizontal-overflow`, await surface().evaluate(e => e.scrollWidth <= e.clientWidth + 1), true);
          await check(`stage:${mode}:${size}:${viewport}:body-no-horizontal-overflow`, await marker(surface(), 'body').evaluate(e => e.scrollWidth <= e.clientWidth + 1), true);
          await check(`stage:${mode}:${size}:${viewport}:same-mode`, await surface().evaluate((e, presentation) => e.classList.contains(presentation === 'dialog' ? 'fui-DialogSurface' : presentation === 'drawer' ? 'fui-OverlayDrawer' : 'fui-PopoverSurface'), mode), true);
          observations.push({ mode, size, viewport, cssViewport, layoutViewport, rect, expectedWidth });
        }
        await capture(`${mode}-${size}-narrow`);
        await page.setViewportSize({ width: 1210, height: 900 }); await settle();
        if (mode === 'popover') {
          const trigger = await rectangle(marker(instance(), 'trigger')), rect = await rectangle(surface());
          await check('stage:popover:anchored-near-trigger', Math.min(Math.abs(rect.y - trigger.bottom), Math.abs(trigger.y - rect.bottom)) <= 16, true);
        }
        await capture(`${mode}-${size}`); await close();
      }
      if (mode !== 'popover') {
        await configure(mode, 'full'); await open();
        for (const [viewport, gap] of [[390, 8], [1000, 10], [1920, 16]]) {
          await page.setViewportSize({ width: viewport, height: 900 }); await settle();
          const rect = await rectangle(surface());
          const cssViewport = await rectangle(page.locator('[data-stage-probe="viewport"]'));
          const cssGap = Math.min(16, Math.max(8, cssViewport.width / 100));
          observations.push({ mode, size: 'full', viewport, gap, cssGap, cssViewport, rect });
          await check(`stage:${mode}:full:${viewport}:width`, approximately(rect.width, cssViewport.width - (mode === 'dialog' ? 2 * cssGap : 0)), true);
          await check(`stage:${mode}:full:${viewport}:height`, approximately(rect.height, cssViewport.height - (mode === 'dialog' ? 2 * cssGap : 0)), true);
          if (mode === 'dialog') {
            await check(`stage:dialog:full:${viewport}:gap`, approximately(rect.x, gap) && approximately(rect.y, gap), true);
            await check(`stage:dialog:full:${viewport}:four-usable-viewport-gaps`, approximately(rect.x, cssGap) && approximately(rect.y, cssGap) && approximately(cssViewport.width - rect.right, cssGap) && approximately(cssViewport.height - rect.bottom, cssGap), true);
            await check(`stage:dialog:full:${viewport}:short-body-fills`, await surface().evaluate(e => {
              const body = e.querySelector('[data-stage-demo="body"]').getBoundingClientRect();
              const header = e.querySelector('[data-stage-demo="header"]').getBoundingClientRect();
              const footer = e.querySelector('[data-stage-demo="footer"]').getBoundingClientRect();
              const gap = parseFloat(getComputedStyle(e.querySelector('[data-stage-demo="form"]')).rowGap);
              return Math.abs(body.y - header.bottom - gap) <= 1 && Math.abs(footer.y - body.bottom - gap) <= 1;
            }), true);
          }
          await capture(`${mode}-full-${viewport}`);
        }
        await close();
      }
    });
    await scenario(`${mode}-scroll`, async () => {
      await configure(mode, mode === 'popover' ? 'large' : 'full', { 'long-body': true }); await open();
      const body = marker(surface(), 'body');
      await check(`stage:${mode}:long-body-scrolls`, await body.evaluate(e => e.scrollHeight > e.clientHeight && getComputedStyle(e).overflowY === 'auto'), true);
      await body.evaluate(e => { e.scrollTop = e.scrollHeight; });
      await check(`stage:${mode}:header-remains-reachable`, await marker(surface(), 'header').isVisible(), true);
      await marker(surface(), 'reset').click();
      await check(`stage:${mode}:footer-action-reachable`, await marker(instance(), 'status').textContent(), 'Reset changes');
      await page.setViewportSize({ width: 390, height: 320 }); await settle();
      await surface().getByRole('button', { name: 'Close A', exact: true }).scrollIntoViewIfNeeded();
      await check(`stage:${mode}:320:header-in-viewport`, await surface().getByRole('button', { name: 'Close A', exact: true }).evaluate(e => { const r = e.getBoundingClientRect(); return r.top >= 0 && r.bottom <= innerHeight + 1; }), true);
      await marker(surface(), 'reset').click();
      await check(`stage:${mode}:320:footer-in-viewport`, await marker(surface(), 'reset').evaluate(e => { const r = e.getBoundingClientRect(); return r.top >= 0 && r.bottom <= innerHeight + 1; }), true);
      await capture(`${mode}-short-viewport`); await close();
    });
  }

  await scenario('popover-arrow', async () => {
    for (const id of ['A', 'B']) {
      await configure('popover', 'medium', {}, id); await open(id);
      const root = surface(id), input = root.getByRole('textbox', { name: `Name ${id}`, exact: true });
      const arrow = () => root.locator(':scope > div').first();
      const mounts = await marker(instance(id), 'mount-count').textContent();
      await input.fill('Arrow draft');
      await check(`stage:arrow:${id}:default-off`, await arrow().evaluate(e => getComputedStyle(e).position === 'absolute'), false);
      await marker(root, 'arrow-toggle').click(); await settle();
      await check(`stage:arrow:${id}:native-arrow`, await arrow().evaluate(e => {
        const s = getComputedStyle(e); return s.position === 'absolute' && s.getPropertyValue('--fui-positioning-arrow-height').trim() !== '';
      }), true);
      await check(`stage:arrow:${id}:surface-does-not-clip`, await root.evaluate(e => getComputedStyle(e).overflow), 'visible');
      await check(`stage:arrow:${id}:no-config-attribute`, await root.getAttribute('witharrow'), null);
      await check(`stage:arrow:${id}:form-preserved`, await input.inputValue(), 'Arrow draft');
      await check(`stage:arrow:${id}:mounts-preserved`, await marker(instance(id), 'mount-count').textContent(), mounts);
      const placement = await root.getAttribute('data-popper-placement');
      await check(`stage:arrow:${id}:native-placement`, /^(top|bottom|left|right)(-start|-end)?$/.test(placement), true);
      await capture(`arrow-${id}-desktop`);
      await page.setViewportSize({ width: 390, height: 320 }); await settle();
      await check(`stage:arrow:${id}:short-viewport-does-not-clip`, await root.evaluate(e => getComputedStyle(e).overflow), 'visible');
      await marker(root, 'arrow-toggle').click(); await settle();
      await check(`stage:arrow:${id}:removed-live`, await arrow().evaluate(e => getComputedStyle(e).position === 'absolute'), false);
      await check(`stage:arrow:${id}:draft-survives-off`, await input.inputValue(), 'Arrow draft');
      await marker(root, 'reset').click(); await close(id);
      await page.setViewportSize({ width: 1210, height: 900 });
    }
  });

  await scenario('sample-layout', async () => {
    for (const [id, mode, viewport] of [
      ['A', 'dialog', 1210], ['A', 'dialog', 320],
      ['A', 'drawer', 1210], ['A', 'drawer', 320],
      ['A', 'popover', 1210], ['A', 'popover', 320],
      ['B', 'dialog', 1210], ['B', 'dialog', 320],
    ]) {
      await page.setViewportSize({ width: viewport, height: 900 });
      await configure(mode, 'small', {}, id); await open(id);
      const layout = await surface(id).evaluate(e => {
        const part = name => e.querySelector(`[data-stage-demo="${name}"]`);
        const header = part('header').getBoundingClientRect(), body = part('body').getBoundingClientRect();
        const footer = part('footer'), bounds = footer.getBoundingClientRect(), css = getComputedStyle(footer);
        const buttons = [...footer.querySelectorAll('button')].map(button => button.getBoundingClientRect());
        const last = buttons[buttons.length - 1];
        return {
          spacing: [getComputedStyle(part('form')).rowGap, css.rowGap, css.columnGap],
          partGaps: [body.top - header.bottom, bounds.top - body.bottom],
          actionLayout: [css.display, css.justifyContent, css.flexWrap],
          endOffset: css.direction === 'rtl' ? last.left - bounds.left : bounds.right - last.right,
          buttonsInside: buttons.every(button => button.left >= bounds.left - 1 && button.right <= bounds.right + 1),
          buttonGaps: buttons.slice(1).map((button, index) => {
            const previous = buttons[index];
            if (Math.abs(button.top - previous.top) > 1) return button.top - previous.bottom;
            return css.direction === 'rtl' ? previous.left - button.right : button.left - previous.right;
          }),
        };
      });
      const name = `stage:sample-layout:${id}:${mode}:${viewport}`;
      await check(`${name}:theme-spacing`, layout.spacing, ['12px', '8px', '8px']);
      await check(`${name}:separated-parts`, layout.partGaps.every(value => approximately(value, 12)), true);
      await check(`${name}:wrapping-end-actions`, layout.actionLayout, ['flex', 'end', 'wrap']);
      await check(`${name}:logical-end`, approximately(layout.endOffset, 0), true);
      await check(`${name}:buttons-fit`, layout.buttonsInside, true);
      await check(`${name}:buttons-separated`, layout.buttonGaps.every(value => value >= 7), true);
      await capture(`sample-layout-${id}-${mode}-${viewport}`); await close(id);
    }
  });

  await scenario('header-semantics', async () => {
    await open();
    const heading = surface().getByRole('heading', { name: 'Profile A', exact: true });
    await check('stage:plain-title', await heading.textContent(), 'Profile A');
    await check('stage:plain-description', await surface().locator('.fui-CardHeader__description').textContent(), 'Edit profile details');
    await check('stage:no-markdown-elements', await marker(surface(), 'header').locator('strong,em,a,code').count(), 0);
    await check('stage:heading-label-association', await surface().getAttribute('aria-labelledby'), await heading.locator('..').getAttribute('id'));
    await check('stage:valid-description-association', await surface().evaluate(e => Boolean(e.getAttribute('aria-describedby') && e.ownerDocument.getElementById(e.getAttribute('aria-describedby'))?.textContent === 'Edit profile details')), true);
    await check('stage:decorative-icon-hidden', await marker(surface(), 'header').locator('.fui-CardHeader__image').getAttribute('aria-hidden'), 'true');
    await check('stage:title-subtitle2', await heading.evaluate(e => { const s = getComputedStyle(e); return [s.fontSize, s.fontWeight, s.lineHeight, s.marginTop, s.marginBottom]; }), ['16px', '600', '22px', '0px', '0px']);
    await check('stage:description-caption1', await surface().locator('.fui-CardHeader__description span').evaluate(e => { const s = getComputedStyle(e); return [s.fontSize, s.lineHeight, s.marginTop, s.marginBottom]; }), ['12px', '16px', '0px', '0px']);
    const card = marker(surface(), 'header').locator('.fui-Card');
    const neutral = () => card.evaluate(e => { const s = getComputedStyle(e), after = getComputedStyle(e, '::after'); return [s.backgroundColor, s.boxShadow, s.paddingTop, s.borderRadius, after.display]; });
    const baseline = await neutral();
    await card.hover(); await check('stage:neutral-header-hover', await neutral(), baseline);
    await page.mouse.down(); await check('stage:neutral-header-pressed', await neutral(), baseline); await page.mouse.up();
    await page.emulateMedia({ forcedColors: 'active' });
    const forced = await neutral();
    await check('stage:neutral-header-forced-colors-transparent', forced[0].endsWith(', 0)'), true);
    await check('stage:neutral-header-forced-colors-structural-styles', forced.slice(1), baseline.slice(1));
    await page.emulateMedia({ forcedColors: 'none' });
    await page.setViewportSize({ width: 390, height: 900 }); await settle();
    await check('stage:header-actions-wrap-at-400-container', await marker(surface(), 'header').evaluate(e => {
      const title = e.querySelector('.fui-CardHeader__header').getBoundingClientRect();
      const action = e.querySelector('.fui-CardHeader__action').getBoundingClientRect();
      return action.top >= title.bottom;
    }), true);
    await marker(surface(), 'header-action').click();
    await check('stage:header-action-event-and-no-dismiss', await marker(surface(), 'action-count').textContent(), 'Header actions: 1');
    await close();
    await configure('dialog', 'small', { jsx: true }); await open();
    await check('stage:jsx-title-preserved', await marker(surface(), 'jsx-title').evaluate(e => e.tagName), 'EM');
    await check('stage:jsx-description-preserved', await marker(surface(), 'jsx-description').locator('strong').textContent(), 'profile');
    await close();
    const ancestor = page.locator('[data-stage-demo="card"]');
    const labelBefore = await ancestor.getAttribute('aria-labelledby');
    const selectionBefore = await page.locator('[data-stage-demo="card-selection"]').textContent();
    await open('B');
    await check('stage:dark-rtl-provider', await surface('B').evaluate(e => [getComputedStyle(e).direction, getComputedStyle(e).getPropertyValue('--colorNeutralForeground1').trim()]), ['rtl', '#ffffff']);
    await check('stage:selectable-card-label-isolated', await ancestor.getAttribute('aria-labelledby'), labelBefore);
    await check('stage:selectable-card-selection-preserved', await page.locator('[data-stage-demo="card-selection"]').textContent(), selectionBefore);
    await capture('dark-rtl'); await close('B');
  });

  for (const mode of ['dialog', 'drawer', 'popover']) {
    await scenario(`${mode}-dismiss-focus`, async () => {
      await configure(mode, 'small'); await open();
      if (mode !== 'popover') {
        const count = await surface().locator('button,input,[tabindex="0"]').count();
        for (const key of ['Tab', 'Shift+Tab']) for (let step = 0; step < count + 2; step++) {
          await page.keyboard.press(key);
          await settle();
          await check(`stage:${mode}:${key}:${step}:modal-focus-contained`, await surface().evaluate(e => e.contains(e.ownerDocument.activeElement)), true);
        }
      } else {
        await page.locator('[data-stage-fixture="outside"]').focus();
        await check('stage:popover:nonmodal-outside-focus', await page.locator('[data-stage-fixture="outside"]').evaluate(e => document.activeElement === e), true);
        if (!(await surface().count())) await open();
      }
      await surface().getByRole('textbox', { name: 'Name A', exact: true }).focus();
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:standard-escape`, await marker(instance(), 'open').textContent(), 'Open: false');
      await check(`stage:${mode}:focus-restored`, await marker(instance(), 'trigger').evaluate(e => document.activeElement === e), true);
      await open(); await page.mouse.click(1, 1); await settle();
      await check(`stage:${mode}:standard-outside-closes`, await marker(instance(), 'open').textContent(), 'Open: false');
      await configure(mode, 'small', { dismiss: 'explicit' }); await open();
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:explicit-escape-refused`, await marker(instance(), 'open').textContent(), 'Open: true');
      await page.mouse.click(1, 1); await settle();
      await check(`stage:${mode}:explicit-outside-refused`, await marker(instance(), 'open').textContent(), 'Open: true');
      await close();
    });
    await scenario(`${mode}-retention`, async () => {
      await configure(mode, 'small', { retention: true }); await open();
      await marker(surface(), 'increment').click(); const mounts = await marker(instance(), 'mount-count').textContent();
      await close();
      const retained = page.locator('[data-stage-demo="surface"][data-stage-instance="A"]');
      if (mode !== 'popover') await check(`stage:${mode}:retained-close-restores-trigger`, await marker(instance(), 'trigger').evaluate(e => document.activeElement === e), true);
      await check(`stage:${mode}:retained-hidden`, await retained.isVisible(), false);
      await check(`stage:${mode}:retained-inert-and-unannounced`, await retained.evaluate(e => e.hasAttribute('hidden') && e.hasAttribute('inert') && e.getAttribute('aria-hidden') === 'true'), true);
      await marker(instance(), 'trigger').focus();
      for (let step = 0; step < 6; step++) {
        await page.keyboard.press('Tab');
        await settle();
        await check(`stage:${mode}:retained-tab-exclusion:${step}`, await retained.evaluate(e => !e.contains(document.activeElement)), true);
        await check(`stage:${mode}:retained-page-keyboard-reachable:${step}`, await page.evaluate(() => document.activeElement !== document.body && document.activeElement.isConnected), true);
      }
      await open();
      await check(`stage:${mode}:retained-counter`, await marker(surface(), 'counter').textContent(), '1');
      await check(`stage:${mode}:retained-mount-count`, await marker(instance(), 'mount-count').textContent(), mounts);
      await close(); await marker(instance(), 'retention').uncheck(); await settle();
      await check(`stage:${mode}:closed-retention-policy-releases-subtree`, await retained.count(), 0);
      await open();
      await marker(surface(), 'increment').click(); await close();
      await check(`stage:${mode}:unmounted-after-close`, await retained.count(), 0);
      await open(); await check(`stage:${mode}:unmount-resets-counter`, await marker(surface(), 'counter').textContent(), '0'); await close();
    });
    await scenario(`${mode}-form-gate`, async () => {
      await configure(mode, 'small', { 'hold-save': true }); await open();
      const name = surface().getByRole('textbox', { name: 'Name A', exact: true });
      await name.fill('Grace'); await surface().getByRole('button', { name: 'Close A', exact: true }).click();
      await marker(surface(), 'discard-prompt').waitFor();
      await check(`stage:${mode}:dirty-refuses-close`, await marker(instance(), 'open').textContent(), 'Open: true');
      await marker(surface(), 'keep-editing').click(); await name.press('Enter');
      await marker(surface(), 'complete-save').waitFor();
      await check(`stage:${mode}:enter-root-gate-once`, await marker(instance(), 'save-count').textContent(), 'Save callbacks: 1');
      await surface().getByRole('button', { name: 'Close A', exact: true }).click();
      await check(`stage:${mode}:pending-refuses-close`, await marker(instance(), 'status').textContent(), 'Close refused: save pending');
      await marker(surface(), 'complete-save').click();
      await marker(instance(), 'status').filter({ hasText: 'Saved Grace' }).waitFor();
      await check(`stage:${mode}:save-clears-dirty`, await marker(surface(), 'dirty').textContent(), 'Dirty: false');
      await name.fill('Changed'); await marker(surface(), 'reset').click();
      await check(`stage:${mode}:reset-restores-saved-baseline`, await name.inputValue(), 'Grace');
      await marker(surface(), 'save').click(); await marker(surface(), 'complete-save').waitFor();
      await check(`stage:${mode}:footer-root-gate-once`, await marker(instance(), 'save-count').textContent(), 'Save callbacks: 2');
      await marker(surface(), 'complete-save').click(); await marker(surface(), 'pending').filter({ hasText: 'Pending: false' }).waitFor();
      await close();
    });
    await scenario(`${mode}-nested-controls`, async () => {
      for (const inline of [false, true]) {
        await configure(mode, 'large', { 'menu-inline': inline }); await open();
        await marker(surface(), 'nested-menu').click();
        await page.getByRole('menuitem', { name: 'Menu action', exact: true }).waitFor();
        await page.keyboard.press('Escape'); await settle();
        await check(`stage:${mode}:menu-${inline}:first-escape-child`, await page.getByRole('menuitem', { name: 'Menu action', exact: true }).isVisible(), false);
        await check(`stage:${mode}:menu-${inline}:stage-stays-open`, await marker(instance(), 'open').textContent(), 'Open: true');
        await marker(surface(), 'nested-menu').click(); await page.getByRole('menuitem', { name: 'Menu action', exact: true }).click();
        await check(`stage:${mode}:menu-${inline}:nested-click-inside`, await marker(instance(), 'open').textContent(), 'Open: true');
        await page.keyboard.press('Escape'); await settle();
        await check(`stage:${mode}:menu-${inline}:second-escape-stage`, await marker(instance(), 'open').textContent(), 'Open: false');
      }
      await configure(mode, 'large'); await open();
      await surface().getByRole('combobox', { name: 'Category A', exact: true }).click();
      await page.getByRole('option', { name: 'Two', exact: true }).waitFor();
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:combobox:first-escape-child`, await page.getByRole('option', { name: 'Two', exact: true }).isVisible(), false);
      await check(`stage:${mode}:combobox:stage-stays-open`, await marker(instance(), 'open').textContent(), 'Open: true');
      await surface().getByRole('combobox', { name: 'Category A', exact: true }).click();
      await page.getByRole('option', { name: 'Two', exact: true }).click();
      await check(`stage:${mode}:combobox:nested-click-inside`, await marker(instance(), 'open').textContent(), 'Open: true');
      await marker(surface(), 'reset').click();
      await marker(surface(), 'date-picker').getByRole('combobox').click();
      const calendar = page.locator('.fui-Calendar'); await calendar.waitFor();
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:datepicker:first-escape-child`, await calendar.isVisible(), false);
      await check(`stage:${mode}:datepicker:stage-stays-open`, await marker(instance(), 'open').textContent(), 'Open: true');
      await surface().getByRole('textbox', { name: 'Name A', exact: true }).focus();
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:datepicker:second-escape-from-stage-control`, await marker(instance(), 'open').textContent(), 'Open: false');
      await open();
      await marker(surface(), 'date-picker').getByRole('combobox').click(); await calendar.waitFor();
      await page.keyboard.press('Escape'); await settle();
      // The installed compat DatePicker consumes Escape on its input even while
      // closed. Preserve that legacy child contract and report it explicitly.
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:datepicker:legacy-closed-input-consumes-Escape`, await marker(instance(), 'open').textContent(), 'Open: true');
      await surface().getByRole('textbox', { name: 'Name A', exact: true }).focus();
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:datepicker:stage-control-dismisses-after-legacy-consumption`, await marker(instance(), 'open').textContent(), 'Open: false');
      await open();
      await marker(surface(), 'date-picker').getByRole('combobox').click(); await calendar.waitFor();
      await calendar.getByRole('button').first().click();
      await check(`stage:${mode}:datepicker:nested-click-inside`, await marker(instance(), 'open').textContent(), 'Open: true');
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:datepicker:portal-escape-closes-child`, await calendar.isVisible(), false);
      await surface().getByRole('textbox', { name: 'Name A', exact: true }).focus();
      await page.keyboard.press('Escape'); await settle();
      await check(`stage:${mode}:datepicker:portal-following-escape-stage`, await marker(instance(), 'open').textContent(), 'Open: false');
    });
  }
  await scenario('independent-providers', async () => {
    await configure('popover', 'large', { dismiss: 'explicit' }); await configure('popover', 'large', { dismiss: 'explicit' }, 'B');
    await open(); await marker(surface(), 'increment').click();
    await open('B');
    await check('stage:independent-instances-open-together', await surface().count() + await surface('B').count(), 2);
    await check('stage:independent-instance-A-counter-retained', await marker(surface(), 'counter').textContent(), '1');
    await check('stage:independent-instance-B-counter', await marker(surface('B'), 'counter').textContent(), '0');
    await check('stage:independent-instance-B-direction', await surface('B').evaluate(e => getComputedStyle(e).direction), 'rtl');
    await close('B');
    await check('stage:independent-instance-B-close-preserves-A', await marker(instance(), 'open').textContent(), 'Open: true');
    await check('stage:independent-instance-A-theme-direction', await surface().evaluate(e => [getComputedStyle(e).direction, getComputedStyle(e).getPropertyValue('--colorNeutralForeground1').trim()]), ['ltr', '#242424']);
    await close();
  });
  await scenario('optional-parts', async () => {
    await configure('dialog', 'small', { description: false, icon: false, close: false, 'footer-toggle': false }); await open();
    await check('stage:optional-description-association-omitted', await surface().evaluate(e => !e.hasAttribute('aria-describedby')), true);
    await check('stage:optional-icon-omitted', await marker(surface(), 'header').locator('.fui-CardHeader__image').count(), 0);
    await check('stage:optional-footer-omitted', await marker(surface(), 'footer').count(), 0);
    await check('stage:optional-X-omitted', await surface().getByRole('button', { name: 'Close A', exact: true }).count(), 0);
    await marker(surface(), 'body-close').click(); await settle();
    await check('stage:optional-visible-escape-route', await marker(instance(), 'open').textContent(), 'Open: false');
  });
  const probe = name => page.locator(`[data-stage-probe="${name}"]`);
  const updateProbe = async patch => { await page.evaluate(value => window.stageProbe.configure(value), patch); await settle(); };
  await scenario('supplemental-jsx-theme-actions', async () => {
    await probe('trigger').click(); await probe('surface').waitFor();
    await check('stage:probe:JSX-refs-preserved', await page.evaluate(() => { const result = window.stageProbe.observe(); return result.titleRef && result.descriptionRef; }), true);
    await check('stage:probe:JSX-authored-styles', await probe('jsx-title').evaluate(e => [getComputedStyle(e).color, getComputedStyle(e).fontSize]), ['rgb(130, 40, 180)', '18px']);
    await check('stage:probe:JSX-description-style', await probe('jsx-description').evaluate(e => [getComputedStyle(e).fontStyle, getComputedStyle(e).letterSpacing]), ['italic', '1px']);
    await probe('jsx-title').click(); await check('stage:probe:JSX-event-preserved', await probe('jsx-clicks').textContent(), '1');
    await probe('increment').click(); const before = await page.evaluate(() => window.stageProbe.observe());
    await updateProbe({ dark: true });
    await check('stage:probe:live-theme', await probe('surface').evaluate(e => getComputedStyle(e).getPropertyValue('--colorNeutralForeground1').trim()), '#ffffff');
    await check('stage:probe:theme-does-not-remount', (await page.evaluate(() => window.stageProbe.observe())).mounts, before.mounts);
    await check('stage:probe:theme-retains-state', await probe('counter').textContent(), '1');
    await updateProbe({ longActions: true }); await page.setViewportSize({ width: 390, height: 900 }); await settle();
    await check('stage:probe:long-actions-no-horizontal-overflow', await probe('surface').evaluate(e => e.scrollWidth <= e.clientWidth + 1), true);
    await capture('long-header-actions');
    for (const name of ['First longer Header action text', 'Second longer Header action text', 'Close probe']) {
      await probe('surface').getByRole('button', { name, exact: true }).click();
      if (name !== 'Close probe') await check(`stage:probe:${name}:no-dismiss`, await probe('open').textContent(), 'true');
    }
    await check('stage:probe:long-actions-close-reachable', await probe('open').textContent(), 'false');
  });
  await scenario('supplemental-event-prevention', async () => {
    await updateProbe({ preventTrigger: true }); await probe('trigger').click(); await settle();
    await check('stage:probe:prevented-trigger-does-not-open', await probe('open').textContent(), 'false');
    await updateProbe({ preventTrigger: false, preventEscape: true }); await probe('trigger').click();
    await probe('surface').getByRole('textbox', { name: 'Probe input' }).focus(); await page.keyboard.press('Escape'); await settle();
    await check('stage:probe:consumer-prevented-Escape', await probe('open').textContent(), 'true');
    await updateProbe({ preventEscape: false });
    if (await probe('surface').count()) { await probe('surface').getByRole('textbox', { name: 'Probe input' }).focus(); await page.keyboard.press('Escape'); }
    await settle(); await check('stage:probe:ordinary-Escape-closes', await probe('open').textContent(), 'false');
  });
  await scenario('supplemental-removed-trigger', async () => {
    await probe('fallback').focus(); await probe('trigger').click(); await probe('surface').waitFor();
    await updateProbe({ trigger: false });
    await probe('surface').getByRole('button', { name: 'Close probe', exact: true }).click(); await settle();
    await check('stage:probe:removed-trigger-stays-disconnected', await probe('trigger').count(), 0);
    await check('stage:probe:removed-trigger-safe-focus', await page.evaluate(() => document.activeElement.isConnected && !document.activeElement.closest('[data-stage-probe="surface"]')), true);
    await probe('fallback').focus(); await page.keyboard.press('Tab');
    await check('stage:probe:removed-trigger-keyboard-page-reachable', await page.evaluate(() => document.activeElement.isConnected && document.activeElement !== document.body), true);
  });
  await scenario('supplemental-iframe-owner', async () => {
    const frame = page.frameLocator('[data-stage-probe="iframe"]');
    const frameProbe = name => frame.locator(`[data-stage-probe="${name}"]`);
    await frameProbe('trigger').click(); await frameProbe('surface').waitFor(); await settle();
    await check('stage:probe:iframe-portal-ownerDocument', await frameProbe('surface').evaluate(e => e.ownerDocument !== window.parent.document && e.ownerDocument === document), true);
    await check('stage:probe:iframe-theme-direction', await frameProbe('surface').evaluate(e => [getComputedStyle(e).direction, getComputedStyle(e).getPropertyValue('--colorNeutralForeground1').trim()]), ['rtl', '#ffffff']);
    await check('stage:probe:iframe-renderer-inserts-owner-styles', await frame.locator('style[data-make-styles-bucket]').count() > 0, true);
    await frameProbe('increment').click();
    await updateProbe({ visible: false });
    await check('stage:probe:removing-main-instance-preserves-iframe', await frameProbe('counter').textContent(), '1');
    await check('stage:probe:iframe-label-still-owned', await frameProbe('surface').evaluate(e => Boolean(document.getElementById(e.getAttribute('aria-labelledby')))), true);
    await frameProbe('surface').getByRole('textbox', { name: 'Probe input' }).focus();
    for (let step = 0; step < 10; step++) {
      await page.keyboard.press(step % 2 ? 'Shift+Tab' : 'Tab');
      await settle();
      await check(`stage:probe:iframe-focus-trap:${step}`, await frameProbe('surface').evaluate(e => e.contains(document.activeElement)), true);
    }
    await capture('iframe-owner');
    await frameProbe('surface').getByRole('button', { name: 'Close iframe probe', exact: true }).click(); await settle();
    await check('stage:probe:iframe-restores-owner-trigger', await frameProbe('trigger').evaluate(e => document.activeElement === e), true);
    await frameProbe('retention').check();
    await frameProbe('trigger').click(); await frameProbe('surface').waitFor();
    await frameProbe('surface').getByRole('button', { name: 'Close iframe probe', exact: true }).click(); await settle();
    await check('stage:probe:retained-iframe-restores-owner-trigger', await frameProbe('trigger').evaluate(e => document.activeElement === e), true);
    await check('stage:probe:retained-iframe-hidden-inert', await frameProbe('surface').evaluate(e => e.hidden && e.hasAttribute('inert') && e.getAttribute('aria-hidden') === 'true'), true);
    await frameProbe('trigger').click(); await frameProbe('surface').waitFor();
    await frameProbe('surface').getByRole('button', { name: 'Close iframe probe', exact: true }).click();
    await page.locator('[data-stage-fixture="outside"]').focus(); await settle();
    await check('stage:probe:retained-iframe-preserves-parent-focus', await page.locator('[data-stage-fixture="outside"]').evaluate(e => document.activeElement === e), true);
    await marker(instance(), 'trigger').click(); await surface().waitFor(); await close();
    await check('stage:probe:iframe-does-not-break-real-sample', await marker(instance(), 'open').textContent(), 'Open: false');
  });
  fs.writeFileSync(path.join(output, `${label}-stage-scenarios.json`), JSON.stringify({ boundary, browserRouting: 'Browser plugin not available', flow: 'Actual StagePanel -> configure/open Stage -> geometry, keyboard, dismissal, retention and Form.Root behavior; supplementary native Stage probes for iframe/theme/JSX/event/trigger contracts', viewportRule: 'Native Fluent scrollbar-gutter:stable reserves classic scrollbar width; independent class-based 100vw/100vh oracle measures usable CSS viewport. Raw physical viewport, CSS viewport and Stage rects are retained.', cases, observations, consoleMessages, limitations: ['Installed compat DatePicker consumes Escape on its own input even when the calendar is closed; focus another Stage control for Escape dismissal. Native child prevention is preserved.'], unverified: ['Physical screen reader and authenticated SharePoint/Teams'] }, null, 2));
}
module.exports = { checkStageBrowser };
