// Standalone real-browser checkpoint; uses an explicitly supplied installed Playwright.
// No @playwright/test runner, network service, tenant or new dependency is needed.
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE_PATH || 'playwright');
const expected = require('../tests/fixtures/sx-browser/expected.json');
const output = path.resolve(process.env.SX_BROWSER_EVIDENCE || '.docs/maintenance/evidence/use-sx');
const label = process.env.SX_BROWSER_RUN_LABEL || 'task-3-browser';
const url = process.env.SX_BROWSER_URL || 'http://127.0.0.1:4317';
const observations = [];
const errors = [];
const consoleMessages = [];
const thumbInteractionCaptures = [];
let browser;
let page;
async function check(name, actual, wanted) {
  const pass = typeof wanted === 'object' ? JSON.stringify(actual) === JSON.stringify(wanted) : actual === wanted;
  observations.push({ name, actual, expected: wanted, pass });
  if (!pass) errors.push(`${name}: expected ${JSON.stringify(wanted)}, got ${JSON.stringify(actual)}`);
}
async function computed(id, property) {
  return page.locator(`#${id}`).evaluate((element, key) => getComputedStyle(element)[key], property);
}
async function scrollbarAxes(id) {
  return page.locator(`#${id}`).evaluate(element => {
    const css = getComputedStyle(element, '::-webkit-scrollbar');
    return { width: css.width, height: css.height };
  });
}
async function exerciseScrollbarThumb(id, axis, name, colors) {
  const target = page.locator(`#${id}`);
  await target.scrollIntoViewIfNeeded();
  await target.evaluate(element => { element.scrollTop = 0; element.scrollLeft = 0; });
  const geometry = await target.evaluate(element => {
    const rect = element.getBoundingClientRect(), css = getComputedStyle(element);
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height,
      left: parseFloat(css.borderLeftWidth), right: parseFloat(css.borderRightWidth),
      top: parseFloat(css.borderTopWidth), bottom: parseFloat(css.borderBottomWidth) };
  });
  const point = axis === 'vertical'
    ? { x: geometry.x + geometry.width - geometry.right - 3, y: geometry.y + geometry.top + 10 }
    : { x: geometry.x + geometry.left + 10, y: geometry.y + geometry.height - geometry.bottom - 3 };
  let currentPoint = { x: 0, y: 0 };
  async function capture(phase, expectedColor) {
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    const filename = `${label}-${name}-${axis}-${phase}.png`;
    await target.screenshot({ path: path.join(output, filename) });
    // CSSOM reports the base pseudo style for native scrollbar widget states.
    // Preserve actual paint evidence for pixel review instead of certifying a
    // hover/pressed color from that unrelated computed base value.
    thumbInteractionCaptures.push({ name, axis, phase, expectedColor, filename,
      pointer: currentPoint, geometry, paintVerification: 'Requires screenshot/pixel review' });
  }
  await page.mouse.move(0, 0);
  await capture('base', colors.base);
  currentPoint = point;
  await page.mouse.move(point.x, point.y);
  await capture('hover', colors.hover);
  await page.mouse.down();
  try {
    await capture('pressed', colors.pressed);
    currentPoint = { x: point.x + (axis === 'horizontal' ? 30 : 0), y: point.y + (axis === 'vertical' ? 15 : 0) };
    await page.mouse.move(currentPoint.x, currentPoint.y, { steps: 4 });
    await capture('drag', colors.pressed);
    await check(`${name}:${axis}:thumb-drag-scrolls`, await target.evaluate((element, direction) =>
      direction === 'vertical' ? element.scrollTop > 0 : element.scrollLeft > 0, axis), true);
  } finally { await page.mouse.up(); }
  await capture('released', colors.hover);
  currentPoint = { x: 0, y: 0 };
  await page.mouse.move(0, 0);
  await capture('leave', colors.base);
  await check(`${name}:${axis}:retains-6px`, await scrollbarAxes(id), { width: '6px', height: '6px' });
}
async function state(id) {
  return page.locator(`#${id}`).evaluate(element => ({
    hover: element.matches(':hover'), active: element.matches(':active'), focusVisible: element.matches(':focus-visible')
  }));
}
async function render(patch) {
  await page.evaluate(value => window.sxFixture.render(value), patch);
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
}
async function snapshot(name) {
  fs.writeFileSync(path.join(output, `${label}-${name}.html`), await page.content());
  await page.screenshot({ path: path.join(output, `${label}-${name}.png`), fullPage: true });
}
async function outlineVisible(id, name) {
  const outline = await page.locator(`#${id}`).evaluate(element => {
    const css = getComputedStyle(element);
    return { width: css.outlineWidth, style: css.outlineStyle, color: css.outlineColor };
  });
  observations.push({ name: `${name}:outline-computed`, actual: outline, expected: 'positive width, visible style/color',
    pass: parseFloat(outline.width) > 0 && outline.style !== 'none' && outline.style !== 'hidden' && outline.color !== 'rgba(0, 0, 0, 0)' });
  if (!observations.at(-1).pass) errors.push(`${name}: native focus outline is absent: ${JSON.stringify(outline)}`);
}
async function runOrder(reverse) {
  await page.goto(`${url}/?reverse=${reverse ? '1' : '0'}`);
  await page.waitForFunction(() => Boolean(window.sxFixture));
  const order = reverse ? 'inverse' : 'forward';
  await check(`${order}:page-title`, await page.title(), 'useSx real-browser checkpoint');
  await check(`${order}:page-url`, new URL(page.url()).origin, url);
  await check(`${order}:meaningful-content`, await page.locator('h1').textContent(), 'useSx real-browser checkpoint');
  await check(`${order}:framework-overlay`, await page.locator('webpack-dev-server-client-overlay, #webpack-dev-server-client-overlay, nextjs-portal').count(), 0);
  await check(`${order}:real-container-type`, await computed('main-container', 'containerType'), 'inline-size');
  await check(`${order}:real-container-name`, await computed('main-container', 'containerName'), 'apvee-sx');
  for (const size of [479, 480, 639, 640, 1023, 1024]) {
    await render({ size });
    await check(`${order}:${size}:direction`, await computed('direction', 'flexDirection'), expected.containerThresholds[size]);
    await check(`${order}:${size}:display`, await computed('direction', 'display'), 'flex');
    const geometry = await page.locator('#direction').evaluate(element => {
      const first = element.querySelector('[data-part="first"]').getBoundingClientRect();
      const second = element.querySelector('[data-part="second"]').getBoundingClientRect();
      return second.top > first.top ? 'column' : second.left > first.left ? 'row' : 'overlap';
    });
    await check(`${order}:${size}:rendered-flex-layout`, geometry, expected.containerThresholds[size]);
    await check(`${order}:${size}:width`, await computed('responsive-width', 'width'), expected.responsiveWidths[size]);
    if (expected.gap[size]) await check(`${order}:${size}:gap`, await computed('responsive-gap', 'rowGap'), expected.gap[size]);
    await check(`${order}:${size}:viewport-unaffected`, await computed('viewport', 'width'), expected.viewportWidth);
    await check(`${order}:${size}:window-unaffected`, await page.evaluate(() => innerWidth), 1210);
  }
  await check(`${order}:nested-nearest-container`, await computed('inner-child', 'width'), expected.nestedWidth);
  await check(`${order}:outer-large-container`, await computed('outer-child', 'width'), '360px');
  await check(`${order}:independent-regions`, [await computed('region-small-child', 'flexDirection'), await computed('region-medium-child', 'flexDirection')], expected.independentRegions);
  await check(`${order}:outside-container-base`, await computed('outside-container', 'width'), '100px');
  await check(`${order}:container-does-not-measure-itself`, await computed('self-container', 'width'), expected.selfContainerWidth);
  await check(`${order}:container-beats-viewport-same-threshold`, await computed('priority', 'width'), expected.queryPriority);
  await check(`${order}:parent-own-child`, await computed('own-child', 'width'), expected.ownChildWidth);
  await check(`${order}:omitted-color-inherits`, await computed('omitted-color', 'color'), expected.inheritedOmittedColor);
  await page.locator('#parent').hover();
  await check(`${order}:hovered-parent-width`, await computed('parent', 'width'), '600px');
  await check(`${order}:hovered-parent-own-child`, await computed('own-child', 'width'), expected.ownChildWidth);
  await check(`${order}:hovered-parent-child-token-isolated`, await computed('own-child', 'color'),
    await page.locator('#own-child').evaluate(element => {
      const oracle = document.createElement('span');
      oracle.style.color = getComputedStyle(element).getPropertyValue('--colorNeutralForeground1');
      element.appendChild(oracle);
      const color = getComputedStyle(oracle).color;
      oracle.remove();
      return color;
    }));
  await check(`${order}:hovered-parent-own-child-private-var`, await page.locator('#own-child').evaluate(element =>
    getComputedStyle(element).getPropertyValue('--apvee-sx-width-container-640').trim()), '');
  await page.mouse.move(0, 0);
  for (const id of ['separate', 'separate-strings']) {
    await check(`${order}:${id}:base`, await computed(id, 'width'), expected.separateSx.base);
    await page.locator(`#${id}`).hover();
    await check(`${order}:${id}:hover`, await computed(id, 'width'), expected.separateSx.hover);
    await page.mouse.move(0, 0);
  }
  await check(`${order}:separate-query:container`, await computed('separate-query', 'width'), expected.separateSx.container);
  await page.locator('#separate-query').hover();
  await check(`${order}:separate-query:container-hover`, await computed('separate-query', 'width'), expected.separateSx.containerHover);
  await page.mouse.move(0, 0);
  for (const [id, key] of [['external-intercalated', 'intercalated'], ['external-later-descriptor', 'laterDescriptor'], ['external-last', 'externalLast'], ['descriptor-last', 'descriptorLast']]) {
    await check(`${order}:${id}`, await computed(id, 'width'), expected.externalClass[key]);
  }
  // Actual native Griffel property classes and sx bindings share each atomic scope.
  // Inactive query/state descriptors must preserve the preceding foreign base.
  for (const size of [639, 640]) {
    await page.mouse.move(0, 0);
    await page.setViewportSize({ width: size, height: 900 });
    await render({ size });
    for (const scope of ['hover', 'viewport', 'container', 'viewport-hover', 'container-hover']) {
      const hasHover = scope.includes('hover');
      const queryActive = scope === 'hover' || size === 640;
      for (const [suffix, activeWidth] of [['sx-last', expected.foreignScoped.sxLast], ['native-last', expected.foreignScoped.nativeLast], ['base-preserved', expected.foreignScoped.sxLast]]) {
        const id = `foreign-${scope}-${suffix}`;
        await page.mouse.move(0, 0);
        await check(`${order}:${size}:${id}:inactive-state`, await computed(id, 'width'),
          !hasHover && queryActive ? activeWidth : expected.foreignScoped.base);
        if (hasHover) {
          await page.locator(`#${id}`).hover({ position: { x: 5, y: 5 } });
          await check(`${order}:${size}:${id}:real-hover`, (await state(id)).hover, true);
          await check(`${order}:${size}:${id}:hover-width`, await computed(id, 'width'),
            queryActive ? activeWidth : expected.foreignScoped.base);
        }
      }
    }
  }
  await page.mouse.move(0, 0);
  await page.setViewportSize({ width: 1210, height: 900 });
  await render({ size: 1024 });
  await page.locator('#scope-removed').scrollIntoViewIfNeeded();
  const removedScopeBox = await page.locator('#scope-removed').boundingBox();
  // Keep the pointer inside both the 333px hovered size and the 111px base size.
  await page.mouse.move(removedScopeBox.x + 5, removedScopeBox.y + 5);
  await check(`${order}:scope-present-hover`, await computed('scope-removed', 'width'), '333px');
  await render({ scopes: false });
  await check(`${order}:removed-scope-still-hovered`, (await state('scope-removed')).hover, true);
  await check(`${order}:removed-scope-width`, await computed('scope-removed', 'width'), expected.removedScopeWidth);
  await page.mouse.move(0, 0);
  await render({ size: 639 });
  await check(`${order}:vertical-639`, await computed('vertical-child', 'flexDirection'), expected.verticalInlineThreshold[639]);
  await render({ size: 640 });
  await check(`${order}:vertical-640`, await computed('vertical-child', 'flexDirection'), expected.verticalInlineThreshold[640]);
  await check(`${order}:vertical-block-width-stays-160`, await computed('vertical-container', 'width'), '160px');
  await snapshot(`${order}-threshold-640`);

  // Actual pointer and keyboard events, with all pseudo-states checked explicitly.
  await page.locator('#state-base').hover();
  await check(`${order}:hover-color`, await computed('state-base', 'backgroundColor'), expected.simultaneousStates.hover);
  await page.mouse.down();
  await check(`${order}:pointer-active-hover`, await state('state-base'), { hover: true, active: true, focusVisible: false });
  await check(`${order}:active-beats-hover`, await computed('state-base', 'backgroundColor'), expected.simultaneousStates.active);
  await page.mouse.up();
  await page.locator('#focus-start').focus();
  await page.keyboard.press('Tab');
  await check(`${order}:keyboard-target`, await page.evaluate(() => document.activeElement.id), 'state-base');
  await check(`${order}:keyboard-focus-color`, await computed('state-base', 'backgroundColor'), expected.simultaneousStates.focusVisible);
  await page.locator('#state-base').hover();
  await page.keyboard.down('Space');
  await check(`${order}:all-three-states`, await state('state-base'), { hover: true, active: true, focusVisible: true });
  await check(`${order}:focus-visible-survives-active-hover`, await computed('state-base', 'backgroundColor'), expected.simultaneousStates.focusVisible);
  await outlineVisible('state-base', `${order}:simultaneous-state-focus`);
  await snapshot(`${order}-simultaneous-states`);
  await page.keyboard.up('Space');
  await page.keyboard.press('Tab');
  await page.locator('#state-query').hover();
  await page.keyboard.down('Space');
  await check(`${order}:query-all-three-states`, await state('state-query'), { hover: true, active: true, focusVisible: true });
  await check(`${order}:query-focus-visible-survives`, await computed('state-query', 'backgroundColor'), expected.simultaneousStates.focusVisibleContainer);
  await outlineVisible('state-query', `${order}:query-focus`);
  await page.keyboard.up('Space');
  await page.keyboard.press('Tab');
  await check(`${order}:keyboard-priority-target`, await page.evaluate(() => document.activeElement.id), 'state-priority');
  for (const size of [479, 480, 640, 1024]) {
    // Clear the retained pointer before resize: a wrapping button's hover width
    // can change its height and otherwise oscillate under the old pointer.
    await page.mouse.move(0, 0);
    await page.setViewportSize({ width: size, height: 900 });
    await render({ size });
    await page.locator('#state-priority').hover({ position: { x: 5, y: 5 } });
    await page.keyboard.down('Space');
    await check(`${order}:${size}:priority-all-three-states`, await state('state-priority'), { hover: true, active: true, focusVisible: true });
    await check(`${order}:${size}:focus-state-responsive-priority`, await computed('state-priority', 'width'), expected.statePriorityWidths[size].focusVisible);
    await outlineVisible('state-priority', `${order}:${size}:priority-focus`);
    await page.keyboard.up('Space');
  }
  // Change keyboard focus away before measuring pointer-only state priorities.
  await page.locator('#focus-start').focus();
  for (const size of [479, 480, 640, 1024]) {
    // Clear the retained pointer before resize: a wrapping button's hover width
    // can change its height and otherwise oscillate under the old pointer.
    await page.mouse.move(0, 0);
    await page.setViewportSize({ width: size, height: 900 });
    await render({ size });
    await page.locator('#state-priority').hover({ position: { x: 5, y: 5 } });
    await check(`${order}:${size}:hover-state-responsive-priority`, await computed('state-priority', 'width'), expected.statePriorityWidths[size].hover);
    await page.mouse.down();
    await check(`${order}:${size}:active-state-responsive-priority`, await computed('state-priority', 'width'), expected.statePriorityWidths[size].active);
    await page.mouse.up();
  }
  await page.setViewportSize({ width: 1210, height: 900 });
  await page.locator('#viewport-state').hover();
  await check(`${order}:viewport-plus-hover-1210`, await computed('viewport-state', 'backgroundColor'), expected.simultaneousStates.active);
  await page.mouse.move(0, 0);
  await page.locator('#focus-start').focus();
  await page.locator('#transition').hover();
  await check(`${order}:transition-enabled-hover`, await computed('transition', 'backgroundColor'), expected.simultaneousStates.hover);
  const pointerBefore = await state('transition');
  await render({ enabled: false });
  await check(`${order}:disabled-native-attribute`, await page.locator('#transition').isDisabled(), true);
  await check(`${order}:disabled-pointer-remains`, (await state('transition')).hover, pointerBefore.hover);
  await check(`${order}:disabled-color`, await computed('transition', 'backgroundColor'), expected.disabledColor);
  await render({ enabled: true, selected: true });
  await check(`${order}:selected-pointer-remains`, (await state('transition')).hover, true);
  await check(`${order}:selected-native-state`, await page.locator('#transition').getAttribute('aria-pressed'), 'true');
  await check(`${order}:selected-color`, await computed('transition', 'backgroundColor'), expected.selectedColor);
  await render({ selected: false });
  await check(`${order}:enabled-hover-restored`, await computed('transition', 'backgroundColor'), expected.simultaneousStates.hover);
  await render({ selected: true });
  await page.locator('#viewport-state').focus();
  await page.keyboard.press('Tab');
  await check(`${order}:selected-keyboard-target`, await page.evaluate(() => document.activeElement.id), 'transition');
  await check(`${order}:selected-keyboard-focus-visible`, (await state('transition')).focusVisible, true);
  await outlineVisible('transition', `${order}:selected-focus`);
  await snapshot(`${order}-selected-focus`);

  for (const [id, widths] of [
    ['foreign-separate-states', expected.separateStates],
    ['foreign-state-sx-last', { base: '100px', hover: '360px', active: '360px', focusVisible: '360px' }],
    ['foreign-state-native-last', { base: '100px', hover: '400px', active: '400px', focusVisible: '400px' }]
  ]) {
    await page.locator('#focus-start').focus();
    await page.mouse.move(0, 0);
    await check(`${order}:${id}:base`, await computed(id, 'width'), widths.base);
    await page.locator(`#${id}`).hover({ position: { x: 5, y: 5 } });
    await check(`${order}:${id}:hover`, await computed(id, 'width'), widths.hover);
    await page.mouse.down();
    await check(`${order}:${id}:pointer-states`, await state(id), { hover: true, active: true, focusVisible: false });
    await check(`${order}:${id}:active`, await computed(id, 'width'), widths.active);
    await page.mouse.up();
    await page.locator('#focus-start').focus();
    await page.keyboard.press('Tab');
    await page.locator(`#${id}`).focus();
    // Focusing can scroll the target after the pointer-only check. Place the
    // pointer inside its current geometry before asserting all three states.
    await page.locator(`#${id}`).hover({ position: { x: 5, y: 5 } });
    await page.keyboard.down('Space');
    await check(`${order}:${id}:simultaneous-states`, await state(id), { hover: true, active: true, focusVisible: true });
    await check(`${order}:${id}:focus-visible`, await computed(id, 'width'), widths.focusVisible);
    await outlineVisible(id, `${order}:${id}:focus-outline`);
    await page.keyboard.up('Space');
  }
  await page.mouse.move(0, 0);

  await render({ dir: 'rtl' });
  for (const [property, value] of Object.entries({ paddingLeft: expected.rtl.paddingLeft, paddingRight: expected.rtl.paddingRight, paddingInlineStart: expected.rtl.logicalStart, direction: expected.rtl.direction })) {
    await check(`${order}:rtl:${property}`, await computed('logical', property), value);
  }
  await check(`${order}:rtl-physical-flipped-left`, await computed('physical', 'paddingLeft'), '0px');
  await check(`${order}:rtl-physical-flipped-right`, await computed('physical', 'paddingRight'), '17px');
  await check(`${order}:sx-does-not-add-inline-style`, await page.locator('.probe[style], button[style]').count(), 0);
  await snapshot(`${order}-rtl`);
  await page.setViewportSize({ width: 639, height: 900 });
  await check(`${order}:viewport-639`, await computed('viewport', 'width'), '120px');
  await page.locator('#viewport-state').hover();
  await check(`${order}:viewport-plus-hover-639`, await computed('viewport-state', 'backgroundColor'), expected.simultaneousStates.hover);
  await page.setViewportSize({ width: 640, height: 900 });
  await check(`${order}:viewport-inclusive-640`, await computed('viewport', 'width'), expected.viewportWidth);
  await check(`${order}:viewport-plus-hover-inclusive-640`, await computed('viewport-state', 'backgroundColor'), expected.simultaneousStates.active);
  await page.setViewportSize({ width: 390, height: 844 });
  await check(`${order}:mobile-viewport`, await computed('viewport', 'width'), '120px');
  await snapshot(`${order}-mobile`);
  await page.setViewportSize({ width: 1210, height: 900 });
}
// Hand-checked floor palette oracles. These do not read the public descriptor factories.
const catalogPalette = {
  light: { canvas: ['#ffffff', '#242424'], alternative: ['#fafafa', '#242424'], subtle: ['transparent', '#424242'],
    transparent: ['transparent'], brand: ['#0f6cbd', '#ffffff'], brandTint: ['#ebf3fc', '#115ea3'],
    inverted: ['#292929', '#ffffff'], success: ['#f1faf1', '#0e700e'], warning: ['#fff9f5', '#bc4b09'],
    danger: ['#fdf3f4', '#b10e1c'], disabled: ['#f0f0f0', '#bdbdbd'] },
  dark: { canvas: ['#292929', '#ffffff'], alternative: ['#1f1f1f', '#ffffff'], subtle: ['transparent', '#d6d6d6'],
    transparent: ['transparent'], brand: ['#115ea3', '#ffffff'], brandTint: ['#082338', '#62abf5'],
    inverted: ['#ffffff', '#242424'], success: ['#052505', '#54b054'], warning: ['#4a1e04', '#faa06b'],
    danger: ['#3b0509', '#dc626d'], disabled: ['#141414', '#5c5c5c'] },
  highContrast: { canvas: ['#000000', '#ffffff'], alternative: ['#000000', '#ffffff'], subtle: ['transparent', '#ffffff'],
    transparent: ['transparent'], brand: ['#ffffff', '#000000'], brandTint: ['#000000', '#ffffff'],
    inverted: ['#000000', '#000000'], success: ['#000000', '#ffffff'], warning: ['#000000', '#ffffff'],
    danger: ['#000000', '#ffffff'], disabled: ['#000000', '#3ff23f'] }
};
const catalogContrasts = [];
function rgb(hex) {
  if (hex === 'transparent') return 'rgba(0, 0, 0, 0)';
  const n = parseInt(hex.slice(1), 16);
  return `rgb(${n >> 16}, ${(n >> 8) & 255}, ${n & 255})`;
}
function luminance(css) {
  const channels = css.match(/[\d.]+/g).slice(0, 3).map(Number).map(value => {
    const s = value / 255;
    return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}
function contrast(first, second) {
  const a = luminance(first), b = luminance(second);
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}
async function runCatalog() {
  await page.goto(url);
  await page.waitForFunction(() => Boolean(window.sxFixture));
  await check('catalog:scrollbar-standard-support', await page.evaluate(() =>
    CSS.supports('scrollbar-width', 'thin') && CSS.supports('scrollbar-color', 'red transparent')), true);
  const vendorScrollbars = await page.evaluate(() => CSS.supports('selector(::-webkit-scrollbar)'));
  if (vendorScrollbars) await check('catalog:thumb-state-selector-support', await page.evaluate(() =>
    ['::-webkit-scrollbar-thumb:hover', '::-webkit-scrollbar-thumb:active', '::-webkit-scrollbar-thumb:hover:active']
      .every(selector => CSS.supports(`selector(${selector})`))), true);
  const styledWidth = vendorScrollbars ? 'auto' : 'thin';
  const stableClasses = {};
  for (const theme of ['light', 'dark', 'highContrast']) {
    await render({ theme, dir: 'ltr' });
    for (const surface of ['canvas', 'alternative']) {
      const underlying = rgb(catalogPalette[theme][surface][0]);
      for (const [name, [background, foreground]] of Object.entries(catalogPalette[theme])) {
        const id = `preset-${surface}-${name}`, color = await computed(id, 'color'), bg = await computed(id, 'backgroundColor');
        await check(`catalog:${theme}:${surface}:${name}:background`, bg, rgb(background));
        await check(`catalog:${theme}:${surface}:${name}:foreground`, color, rgb(foreground || catalogPalette[theme][surface][1]));
        const className = await page.locator(`#${id}`).getAttribute('class');
        if (theme === 'light') stableClasses[id] = className;
        else await check(`catalog:${theme}:${surface}:${name}:theme-independent-class`, className, stableClasses[id]);
        const ratio = contrast(color, background === 'transparent' ? underlying : bg);
        const unsupported = theme === 'highContrast' && name === 'inverted';
        const disabled = name === 'disabled';
        catalogContrasts.push({ theme, surface, preset: name, foreground: color, background: bg, underlying,
          ratio: Number(ratio.toFixed(2)), supportedForActiveText: !unsupported && !disabled,
          limitation: unsupported ? 'Unsupported inverted pair in floor high-contrast theme: 1:1.'
            : disabled ? 'Disabled presentation only; does not disable the element.' : null });
        if (unsupported) await check(`catalog:${theme}:${surface}:${name}:known-unsupported-ratio`, ratio, 1);
        else if (!disabled) await check(`catalog:${theme}:${surface}:${name}:verified-text-ratio-at-least-4.5`, ratio >= 4.5, true);
      }
      const scrollbarId = `scrollbar-${surface}`;
      await check(`catalog:${theme}:${surface}:scrollbar-width`, await computed(scrollbarId, 'scrollbarWidth'), styledWidth);
      const thumb = rgb({ light: '#616161', dark: '#adadad', highContrast: '#ffffff' }[theme]);
      await check(`catalog:${theme}:${surface}:scrollbar-color`, await computed(scrollbarId, 'scrollbarColor'), vendorScrollbars ? 'auto' : `${thumb} rgba(0, 0, 0, 0)`);
      if (vendorScrollbars) {
        await check(`catalog:${theme}:${surface}:scrollbar-both-axes-6px`, await scrollbarAxes(scrollbarId), { width: '6px', height: '6px' });
        await check(`catalog:${theme}:${surface}:scrollbar-thumb`, await page.locator(`#${scrollbarId}`).evaluate(element => getComputedStyle(element, '::-webkit-scrollbar-thumb').backgroundColor), thumb);
        await check(`catalog:${theme}:${surface}:scrollbar-rounded-thumb`, await page.locator(`#${scrollbarId}`).evaluate(element => getComputedStyle(element, '::-webkit-scrollbar-thumb').borderTopLeftRadius), '3px');
      }
      await check(`catalog:${theme}:${surface}:scrollbar-overflow-explicit`, await computed(scrollbarId, 'overflowY'), 'auto');
      await check(`catalog:${theme}:${surface}:scrollbar-actual-overflow`, await page.locator(`#${scrollbarId}`).evaluate(element => element.scrollHeight > element.clientHeight), true);
      await check(`catalog:${theme}:${surface}:scrollbar-horizontal-overflow`, await page.locator(`#${scrollbarId}`).evaluate(element => element.scrollWidth > element.clientWidth), true);
      await page.locator(`#${scrollbarId}`).evaluate(element => { element.scrollTop = 40; });
      await check(`catalog:${theme}:${surface}:scrollbar-scrolls`, await page.locator(`#${scrollbarId}`).evaluate(element => element.scrollTop), 40);
      await page.locator(`#${scrollbarId}`).evaluate(element => { element.scrollLeft = 40; });
      await check(`catalog:${theme}:${surface}:scrollbar-scrolls-horizontally`, await page.locator(`#${scrollbarId}`).evaluate(element => element.scrollLeft), 40);
      const ratio = contrast(thumb, underlying);
      catalogContrasts.push({ theme, surface, scrollbarThumb: thumb, underlying, ratio: Number(ratio.toFixed(2)), limitation: 'Transparent track depends on the actual underlying surface.' });
      await check(`catalog:${theme}:${surface}:verified-scrollbar-ratio-at-least-3`, ratio >= 3, true);
      if (vendorScrollbars) {
        const colors = {
          base: thumb,
          hover: rgb({ light: '#575757', dark: '#bdbdbd', highContrast: '#1aebff' }[theme]),
          pressed: rgb({ light: '#4d4d4d', dark: '#b3b3b3', highContrast: '#1aebff' }[theme])
        };
        for (const axis of ['vertical', 'horizontal']) await exerciseScrollbarThumb(scrollbarId, axis, `catalog-${theme}-${surface}`, colors);
      }
    }
    await check(`catalog:${theme}:typography-inherits`, await computed('catalog-typography-inherits', 'color'), 'rgb(11, 22, 33)');
    await check(`catalog:${theme}:transparent-inherits`, await computed('catalog-transparent-inherits', 'color'), 'rgb(11, 22, 33)');
    await check(`catalog:${theme}:preset-foreground-override`, await computed('catalog-preset-override', 'color'), rgb(catalogPalette[theme].subtle[1]));
    await check(`catalog:${theme}:preset-override-background-retained`, await computed('catalog-preset-override', 'backgroundColor'), rgb(catalogPalette[theme].canvas[0]));
    await check(`catalog:${theme}:preset-reverse-foreground`, await computed('catalog-preset-reverse', 'color'), rgb(catalogPalette[theme].canvas[1]));
    await check(`catalog:${theme}:typography-size`, await computed('catalog-typography-body1', 'fontSize'), '14px');
    await check(`catalog:${theme}:typography-weight`, await computed('catalog-typography-body1', 'fontWeight'), '400');
    await check(`catalog:${theme}:typography-lineheight`, await computed('catalog-typography-body1', 'lineHeight'), '20px');
    await check(`catalog:${theme}:scrollbar-only-native-properties`, await page.locator('#scrollbar-only').evaluate(element => {
      const css = getComputedStyle(element);
      return { overflowX: css.overflowX, overflowY: css.overflowY, scrollbarGutter: css.scrollbarGutter,
        overscrollBehavior: css.overscrollBehavior, scrollBehavior: css.scrollBehavior, forcedColorAdjust: css.forcedColorAdjust };
    }), { overflowX: 'visible', overflowY: 'visible', scrollbarGutter: 'auto', overscrollBehavior: 'auto', scrollBehavior: 'auto', forcedColorAdjust: 'auto' });
    await check(`catalog:${theme}:no-inline-style`, await page.locator('#catalog [style]').count(), 0);
    await page.mouse.move(0, 0);
    await check(`catalog:${theme}:scoped-scrollbar-inactive-query-width`, await computed('scrollbar-scoped-inactive', 'scrollbarWidth'), 'auto');
    await check(`catalog:${theme}:scoped-scrollbar-inactive-query-color`, await computed('scrollbar-scoped-inactive', 'scrollbarColor'), 'auto');
    await check(`catalog:${theme}:scoped-scrollbar-active-query-width`, await computed('scrollbar-scoped-query', 'scrollbarWidth'), styledWidth);
    await check(`catalog:${theme}:scoped-scrollbar-inactive-hover-width`, await computed('scrollbar-scoped-query-hover', 'scrollbarWidth'), 'auto');
    await page.locator('#scrollbar-scoped-query-hover').hover({ position: { x: 5, y: 5 } });
    await check(`catalog:${theme}:scoped-scrollbar-real-hover`, (await state('scrollbar-scoped-query-hover')).hover, true);
    await check(`catalog:${theme}:scoped-scrollbar-active-hover-width`, await computed('scrollbar-scoped-query-hover', 'scrollbarWidth'), styledWidth);
    if (vendorScrollbars) {
      await check(`catalog:${theme}:inactive-query-has-native-axes-under-styled-parent`, await scrollbarAxes('scrollbar-scoped-inactive'), { width: 'auto', height: 'auto' });
      await check(`catalog:${theme}:active-query-has-6px-axes`, await scrollbarAxes('scrollbar-scoped-query'), { width: '6px', height: '6px' });
      await check(`catalog:${theme}:active-hover-has-6px-axes`, await scrollbarAxes('scrollbar-scoped-query-hover'), { width: '6px', height: '6px' });
    }
    await page.mouse.move(0, 0);
    if (vendorScrollbars) await check(`catalog:${theme}:inactive-hover-has-native-axes`, await scrollbarAxes('scrollbar-scoped-query-hover'), { width: 'auto', height: 'auto' });
    await render({ scrollbarEnabled: false });
    await check(`catalog:${theme}:removed-scrollbar-width`, await computed('scrollbar-removable', 'scrollbarWidth'), 'auto');
    await check(`catalog:${theme}:removed-scrollbar-color`, await computed('scrollbar-removable', 'scrollbarColor'), 'auto');
    if (vendorScrollbars) await check(`catalog:${theme}:removed-scrollbar-has-native-axes`, await scrollbarAxes('scrollbar-removable'), { width: 'auto', height: 'auto' });
    await render({ scrollbarEnabled: true });
    if (vendorScrollbars) await check(`catalog:${theme}:restored-scrollbar-has-6px-axes`, await scrollbarAxes('scrollbar-removable'), { width: '6px', height: '6px' });
    for (const scope of ['width', 'color', 'hover']) {
      for (const order of ['native-last', 'recipe-last']) {
        for (const composition of ['segmented', 'merged']) {
          const id = `scrollbar-interop-${scope}-${order}-${composition}`;
          if (scope === 'hover') await page.locator(`#${id}`).hover({ position: { x: 5, y: 5 } });
          const nativeLast = order === 'native-last';
          await check(`catalog:${theme}:${id}:width`, await computed(id, 'scrollbarWidth'),
            nativeLast && scope !== 'color' ? 'none' : styledWidth);
          await check(`catalog:${theme}:${id}:color`, await computed(id, 'scrollbarColor'),
            nativeLast && scope !== 'width' ? 'rgb(255, 0, 0) rgba(0, 0, 0, 0)' : vendorScrollbars ? 'auto' : `${rgb({ light: '#616161', dark: '#adadad', highContrast: '#ffffff' }[theme])} rgba(0, 0, 0, 0)`);
          if (vendorScrollbars && !nativeLast) await check(`catalog:${theme}:${id}:recipe-last-6px-axes`, await scrollbarAxes(id), { width: '6px', height: '6px' });
          await page.mouse.move(0, 0);
          if (vendorScrollbars && scope === 'hover') {
            await check(`catalog:${theme}:${id}:inactive-hover-native-width`, await computed(id, 'scrollbarWidth'), 'auto');
            await check(`catalog:${theme}:${id}:inactive-hover-native-axes`, await scrollbarAxes(id), { width: 'auto', height: 'auto' });
          }
        }
      }
    }
    await page.locator('#catalog').scrollIntoViewIfNeeded();
  }
  // High-contrast Fluent palette and browser forced-colors are separate mechanisms.
  for (const theme of ['light', 'dark', 'highContrast']) {
    await render({ theme });
    await page.emulateMedia({ forcedColors: 'active' });
    await check(`catalog:${theme}:forced-colors-is-active`, await page.evaluate(() => matchMedia('(forced-colors: active)').matches), true);
    for (const surface of ['canvas', 'alternative']) {
      await check(`catalog:${theme}:${surface}:forced-scrollbar-color`, await computed(`scrollbar-${surface}`, 'scrollbarColor'), 'auto');
      await check(`catalog:${theme}:${surface}:forced-scrollbar-width`, await computed(`scrollbar-${surface}`, 'scrollbarWidth'), 'thin');
      await check(`catalog:${theme}:${surface}:forced-colors-native-adjust`, await computed(`scrollbar-${surface}`, 'forcedColorAdjust'), 'auto');
      if (vendorScrollbars) await check(`catalog:${theme}:${surface}:forced-colors-vendor-axes-inactive`, await scrollbarAxes(`scrollbar-${surface}`), { width: 'auto', height: 'auto' });
    }
    await check(`catalog:${theme}:forced-scoped-query-scrollbar-color`, await computed('scrollbar-scoped-query', 'scrollbarColor'), 'auto');
    await check(`catalog:${theme}:forced-scoped-query-scrollbar-width`, await computed('scrollbar-scoped-query', 'scrollbarWidth'), 'thin');
    await page.locator('#scrollbar-scoped-query-hover').hover({ position: { x: 5, y: 5 } });
    await check(`catalog:${theme}:forced-scoped-hover-is-active`, (await state('scrollbar-scoped-query-hover')).hover, true);
    await check(`catalog:${theme}:forced-scoped-hover-scrollbar-color`, await computed('scrollbar-scoped-query-hover', 'scrollbarColor'), 'auto');
    await check(`catalog:${theme}:forced-scoped-hover-scrollbar-width`, await computed('scrollbar-scoped-query-hover', 'scrollbarWidth'), 'thin');
    await page.mouse.move(0, 0);
    await page.locator('#catalog').scrollIntoViewIfNeeded();
    await snapshot(`catalog-${theme}-forced-colors`);
    await page.locator('#catalog').screenshot({ path: path.join(output, `${label}-catalog-${theme}-forced-colors-detail.png`) });
    await page.emulateMedia({ forcedColors: 'none' });
    await check(`catalog:${theme}:forced-colors-deactivated`, await page.evaluate(() => matchMedia('(forced-colors: active)').matches), false);
  }
  // Full-page and oversized element captures temporarily resize Chromium's
  // viewport. On macOS that
  // can leave following native scrollbar widgets without their gutter/hit area
  // despite correct computed pseudo styles. Preserve these overview captures
  // after all actual thumb interactions, so they cannot alter later test input.
  for (const theme of ['light', 'dark', 'highContrast']) {
    await render({ theme });
    await page.locator('#catalog').scrollIntoViewIfNeeded();
    await snapshot(`catalog-${theme}`);
    await page.locator('#catalog').screenshot({ path: path.join(output, `${label}-catalog-${theme}-detail.png`) });
  }
}

(async () => {
  fs.mkdirSync(output, { recursive: true });
  browser = await chromium.launch({ headless: true,
    ...(process.env.SX_BROWSER_EXECUTABLE ? { executablePath: process.env.SX_BROWSER_EXECUTABLE } : {}),
    ignoreDefaultArgs: ['--hide-scrollbars'] });
  page = await browser.newPage({ viewport: { width: 1210, height: 900 } });
  page.on('console', message => { if (['error', 'warning'].includes(message.type())) consoleMessages.push({ type: message.type(), text: message.text() }); });
  page.on('pageerror', error => errors.push(`Uncaught browser error: ${error.message}`));
  const version = browser.version();
  try {
    await runOrder(false);
    await runOrder(true);
    await runCatalog();
    await check('console-health', consoleMessages, []);
  } catch (error) { errors.push(error.stack); }
  const report = { version, node: process.version, url, viewport: { width: 1210, height: 900 },
    module: process.env.PLAYWRIGHT_MODULE_PATH || 'playwright', executable: process.env.SX_BROWSER_EXECUTABLE || 'Playwright default',
    localBrowser: true, sharePointHost: 'NOT EXECUTED', observations, catalogContrasts, thumbInteractionCaptures, consoleMessages, errors };
  fs.writeFileSync(path.join(output, `${label}.json`), JSON.stringify(report, null, 2));
  console.log(JSON.stringify({ version, checks: observations.length, failures: errors.length, errors }, null, 2));
  await browser.close();
  assert.equal(errors.length, 0, 'Real-browser checkpoint failed; see evidence JSON.');
})().catch(async error => { console.error(error); if (browser) await browser.close(); process.exitCode = 1; });
