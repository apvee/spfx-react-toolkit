const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const playwright = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const mode = process.argv[2] || 'source';
if (!['source', 'compiled'].includes(mode)) throw new Error('Expected source or compiled');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, `temp/migrated-slots-${mode}`);
const evidence = path.join(root, '.docs/maintenance/evidence/slot-migration');
fs.mkdirSync(evidence, {recursive:true});
const checks = [];
const errors = [];
const peerWarnings = [];
const colorWarning = `The keyword 'slider-vertical' specified to an 'appearance' property is not standardized. It will be removed in the future. Use <input type=range style="writing-mode: vertical-lr; direction: rtl"> instead.`;
function check(name, condition) { assert.ok(condition, name); checks.push(name); }
(async () => {
  const server = http.createServer((request, response) => {
    if (new URL(request.url, 'http://localhost').pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
    const filename = path.join(fixture, new URL(request.url, 'http://localhost').pathname === '/' ? 'index.html' : new URL(request.url, 'http://localhost').pathname.slice(1));
    if (!filename.startsWith(fixture)) { response.writeHead(403).end(); return; }
    fs.readFile(filename, (error, data) => { if (error) response.writeHead(404).end(); else { response.setHeader('Content-Type', filename.endsWith('.js') ? 'text/javascript' : 'text/html'); response.end(data); } });
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  const browser = await playwright.chromium.launch({ executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
    page.on('pageerror', error => errors.push(String(error)));
    page.on('console', message => { if (message.type() === 'warning' && message.text() === colorWarning) peerWarnings.push(message.text());
      else if (['warning', 'error'].includes(message.type())) errors.push(message.text()); });
    const d = (family, value) => page.locator(`[data-${family}-demo="${value}"]`);
    async function setDiagnostics(expanded) {
      const disclosure = page.getByRole('button', { name: 'Advanced diagnostics', exact: true });
      if (await disclosure.getAttribute('aria-expanded') !== String(expanded)) await disclosure.click();
    }
    async function open(panel) {
      await page.goto(`${base}/?panel=${panel}`);
      await page.locator('#root').waitFor();
      await page.waitForTimeout(150);
      check(`${panel}: identity`, await page.title() === 'Migrated slots sample QA' && page.url().includes(`panel=${panel}`));
      check(`${panel}: meaningful content`, (await page.locator('#root').innerText()).length > 100);
      check(`${panel}: no framework overlay`, await page.locator('vite-error-overlay, nextjs-portal').count() === 0);
      await setDiagnostics(true);
    }
    await open('groups');
    for (const type of ['settings', 'accordion']) {
      check(`groups: ${type} default surface is filled`, await d('groups', type).locator(':scope > .fui-Card').last().evaluate(node => getComputedStyle(node).backgroundColor !== 'rgba(0, 0, 0, 0)'));
      check(`groups: ${type} default surface has a shadow`, await d('groups', type).locator(':scope > .fui-Card').last().evaluate(node => getComputedStyle(node).boxShadow !== 'none'));
    }
    check('groups: direction provider is transparent and matches the frame width', await d('groups','frame').evaluate(node => {
      const parent = node.parentElement;
      return getComputedStyle(parent).backgroundColor === 'rgba(0, 0, 0, 0)' && Math.abs(parent.getBoundingClientRect().width - node.getBoundingClientRect().width) < .1;
    }));
    for (const [type, appearanceMarker, inputMarker] of [['settings','appearance','email'], ['accordion','accordion-appearance','note']]) {
      await d('groups', inputMarker).fill('Preserved surface draft');
      await d('groups', inputMarker).evaluate(node => { window.groupSurfaceInput = node; });
      for (const appearance of ['filled', 'alternate', 'outline']) {
        await d('groups', appearanceMarker).selectOption(appearance);
        const surface = await d('groups', type).locator(':scope > .fui-Card').last().evaluate(node => ({ background: getComputedStyle(node).backgroundColor, shadow: getComputedStyle(node).boxShadow }));
        check(`groups: ${type} ${appearance} surface`, appearance === 'outline' ? surface.background === 'rgba(0, 0, 0, 0)' && surface.shadow === 'none' : surface.background !== 'rgba(0, 0, 0, 0)' && surface.shadow !== 'none');
        check(`groups: ${type} ${appearance} retains caller input`, await d('groups', inputMarker).evaluate(node => node === window.groupSurfaceInput && node.value === 'Preserved surface draft'));
      }
    }
    for (const viewport of [[1280,900], [390,844]]) {
      await page.setViewportSize({width:viewport[0],height:viewport[1]});
      for (const rtl of [false,true]) {
        for (const [frameMarker, widthMarker, directionMarker] of [['frame','resize','direction'], ['accordion-frame','accordion-width','accordion-direction']]) {
          await d('groups', directionMarker).setChecked(rtl);
          for (const requested of [320,500,639,640,900]) {
            await d('groups', widthMarker).selectOption(String(requested));
            const fits = await d('groups', frameMarker).evaluate((node, requested) => {
              const provider = node.parentElement, parent = provider.parentElement.parentElement, style = getComputedStyle(parent);
              const available = parent.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
              const actual = node.getBoundingClientRect().width;
              return Math.abs(actual - Math.min(requested, available)) < .1 && Math.abs(actual - provider.getBoundingClientRect().width) < .1
                && getComputedStyle(provider).backgroundColor === 'rgba(0, 0, 0, 0)' && provider.scrollWidth <= provider.clientWidth + 1;
            }, requested);
            check(`groups: ${frameMarker} ${viewport[0]} ${rtl ? 'RTL' : 'LTR'} requested ${requested} matches bounded provider`, fits);
          }
          await d('groups', widthMarker).selectOption('640'); await d('groups', directionMarker).uncheck();
        }
      }
    }
    await page.setViewportSize({width:1280,height:900});
    await d('groups', 'inspect').click();
    check('groups: root and title refs', (await d('groups', 'observations').innerText()).includes('settings=div, accordion=div, visible headings=1, slot=span'));
    check('groups: native heading', await d('groups', 'title-slot').evaluate(node => node.closest('h2')?.textContent === 'Notification preferences'));
    for (const [type, size] of [['object',24], ['callback',20]]) {
      const bounds = await d('groups', `${type}-icon`).evaluate(node => { const r = node.getBoundingClientRect(); return [r.width,r.height]; });
      check(`groups: ${type} authored 80px SVG bounded to ${size}px`, bounds.every(value => Math.abs(value - size) < .1));
    }
    await d('groups','action-slot').getByRole('button').click();
    check('groups: action callback', (await d('groups','observations').innerText()).includes('Actions: 1'));
    const profile = page.getByRole('button', { name: 'Profile details Your account preferences.' });
    check('groups: trigger native heading', await profile.evaluate(node => !!node.closest('h3')));
    await d('groups','note').fill('Persisted slot QA');
    await profile.click();
    await page.waitForFunction(() => document.querySelector('[data-groups-demo="note"]').closest('.fui-AccordionPanel').hasAttribute('inert'));
    check('groups: closed retained panel inert and collapsed', await profile.getAttribute('aria-expanded') === 'false' && await d('groups','note').evaluate(node => node.closest('.fui-AccordionPanel').hasAttribute('inert')));
    await profile.click();
    check('groups: close/reopen retains input', await d('groups','note').inputValue() === 'Persisted slot QA');
    await d('groups','multiple').check();
    const advanced = page.getByRole('button', { name: 'Advanced preferences', exact: true });
    await advanced.click();
    check('groups: native multiple triggers show Both in selector', await d('groups','open-items').inputValue() === 'profile,advanced' && await profile.getAttribute('aria-expanded') === 'true' && await advanced.getAttribute('aria-expanded') === 'true');
    await profile.click();
    check('groups: closing Profile shows Advanced in selector', await d('groups','open-items').inputValue() === 'advanced' && await advanced.getAttribute('aria-expanded') === 'true');
    await advanced.click();
    check('groups: closing Advanced shows None in selector', await d('groups','open-items').inputValue() === '');
    await profile.click();
    check('groups: opening Profile shows Profile in selector', await d('groups','open-items').inputValue() === 'profile');
    await d('groups','unmount').check(); await profile.click();
    await d('groups','note').waitFor({state:'detached'});
    check('groups: unmount removes closed input', await d('groups','note').count() === 0);
    await profile.click();
    check('groups: unmount/reopen resets input', await d('groups','note').inputValue() === 'Retained note');
    await page.waitForTimeout(350);
    await page.screenshot({path:path.join(evidence,`migrated-${mode}-groups.png`),fullPage:true});
    await d('groups','title').fill('Edited notification preferences');
    await d('groups','description').fill('Edited delivery guidance');
    check('groups: editable heading and description update preview', (await d('groups','settings').innerText()).includes('Edited notification preferences') && (await d('groups','settings').innerText()).includes('Edited delivery guidance'));
    await d('groups','hidden').check();
    check('groups: hidden preview leaves configuration reachable', !await d('groups','settings').isVisible() && await d('groups','hidden').isVisible());
    await d('groups','hidden').uncheck();
    await d('groups','email').fill('qa@example.com');
    await d('groups','email').focus();
    await page.evaluate(() => window.demoFixtureHost.update({theme:'dark'}));
    check('groups: host update retains configuration value and focus', await d('groups','email').inputValue() === 'qa@example.com' && await d('groups','email').evaluate(node => node === document.activeElement));
    await d('groups','resize').selectOption('320'); await page.evaluate(() => window.demoFixtureHost.update({theme:'dark'})); await d('groups','direction').check();
    await page.setViewportSize({width:390,height:844});
    await page.waitForTimeout(350);
    await page.screenshot({path:path.join(evidence,`migrated-${mode}-groups-mobile.png`),fullPage:true});
    await page.setViewportSize({width:1280,height:900});
    await open('state');
    await d('state-message','title').fill('Edited state title'); await d('state-message','description').fill('Edited **state guidance**');
    check('state: editable content updates preview', (await d('state-message','preview').innerText()).includes('Edited state title') && (await d('state-message','preview').innerText()).includes('Edited state guidance'));
    await d('state-message','slots').check(); await d('state-message','heading').selectOption('4');
    await d('state-message','inspect').click();
    check('state: root/slot refs and h4', (await d('state-message','observations').innerText()).includes('div, heading=h4, slot=span'));
    check('state: native h4 title literal', await d('state-message','title-slot').evaluate(node => node.closest('h4')?.textContent === '**Literal slot title**'));
    check('state: callback text literal', (await d('state-message','preview').innerText()).includes('Callback keeps **literal markers**.'));
    await d('state-message','preview').getByRole('button',{name:'Try again'}).click();
    check('state: label object action callback', (await d('state-message','observations').innerText()).includes('Actions: 1'));
    check('state: disabled action', await d('state-message','preview').getByRole('button',{name:'Unavailable'}).isDisabled());
    await d('state-message','role').check(); await d('state-message','inspect').click();
    check('state: semantic alert retained', (await d('state-message','observations').innerText()).includes('role=alert'));
    await page.screenshot({path:path.join(evidence,`migrated-${mode}-state.png`),fullPage:true});
    await open('relative');
    await d('relative-time','fallback-slot').click(); await d('relative-time','inspect-slot').click();
    check('relative: fallback slot ref and event', (await page.locator('#root').innerText()).includes('Fallback element: span. Fallback clicks: 1.'));
    await d('relative-time','preview').click(); await d('relative-time','inspect-ref').click();
    check('relative: time root ref and click', (await d('relative-time','observations').innerText()).includes('Ref element: time. Preview clicks: 1.'));
    check('relative: valid time datetime', await d('relative-time','preview').getAttribute('datetime') === '2026-01-31T12:00:00.000Z');
    await d('relative-time','fallback').fill('QA fallback');
    check('relative: fallback callback updates', await d('relative-time','fallback-slot').innerText() === 'QA fallback');
    await page.screenshot({path:path.join(evidence,`migrated-${mode}-relative.png`),fullPage:true});
    await open('formatted');
    await d('formatted-text','text').fill('**Editable text**');
    check('formatted: markdown text edits preview', await d('formatted-text','preview').innerText() === 'Editable text');
    await d('formatted-text','markdown').uncheck();
    check('formatted: markdown toggle changes actual preview', await d('formatted-text','preview').innerText() === '**Editable text**');
    await d('formatted-text','as').selectOption('h4'); await d('formatted-text','inspect-ref').click();
    check('formatted: native select updates root ref', (await d('formatted-text','observations').innerText()).includes('Ref element: h4'));
    const widthField = d('formatted-text','width');
    await widthField.fill('');
    check('formatted: cleared width keeps an editable empty draft', await widthField.inputValue() === '');
    for (const [digit, draft] of [['6', '6'], ['4', '64'], ['0', '640']]) {
      await widthField.pressSequentially(digit);
      check(`formatted: incremental width draft ${draft} is retained`, await widthField.inputValue() === draft);
      if (draft !== '640') check(`formatted: incomplete width ${draft} retains last valid preview`, await d('formatted-text','preview').evaluate(node => Math.abs(node.parentElement.getBoundingClientRect().width - 320) < 3));
    }
    check('formatted: final typed width updates preview', await d('formatted-text','preview').evaluate(node => Math.abs(node.parentElement.getBoundingClientRect().width - 640) < 3));
    await widthField.fill('900');
    check('formatted: out-of-range draft preserves last valid width', await widthField.inputValue() === '900' && await d('formatted-text','preview').evaluate(node => Math.abs(node.parentElement.getBoundingClientRect().width - 640) < 3));
    await widthField.press('Tab');
    check('formatted: invalid draft resets to last valid width on blur', await widthField.inputValue() === '640');
    await widthField.focus(); await widthField.press('ControlOrMeta+A'); await widthField.pressSequentially('200');
    check('formatted: select-all typing replaces width draft', await widthField.inputValue() === '200');
    check('formatted: editable width changes frame', await d('formatted-text','preview').evaluate(node => Math.abs(node.parentElement.getBoundingClientRect().width - 200) < 3));
    await setDiagnostics(false);
    await page.waitForTimeout(400);
    check('formatted: diagnostics collapse retains mixed JSX and native inert', await d('formatted-text','mixed-jsx').count() === 1 && await d('formatted-text','mixed-jsx').evaluate(node => node.closest('.fui-AccordionPanel').getAttribute('inert') === ''));
    await setDiagnostics(true);
    await open('skeleton');
    await d('auto-skeleton','focus-target').fill('Retained Fluent input note');
    await d('auto-skeleton','focus-load').click();
    await page.waitForFunction(() => document.querySelector('[data-auto-skeleton-demo="preview"]').getAttribute('aria-busy') === 'true');
    check('skeleton: loading creates measured geometry', await d('auto-skeleton','preview').locator('[data-auto-skeleton-shape]').count() > 0);
    await d('auto-skeleton','loading').uncheck();
    check('skeleton: loading retains note and single mount', await d('auto-skeleton','focus-target').inputValue() === 'Retained Fluent input note' && (await d('auto-skeleton','observations').innerText()).includes('Content mounts: 1'));
    await d('auto-skeleton','loading').check();
    const beforeAnimation = await d('auto-skeleton','preview').locator('.fui-SkeletonItem').first().evaluate(node => getComputedStyle(node, '::after').animationName);
    await d('auto-skeleton','animation').selectOption('pulse');
    check('skeleton: animation select updates actual skeleton', (await d('auto-skeleton','preview').locator('.fui-SkeletonItem').first().evaluate(node => getComputedStyle(node, '::after').animationName)) !== beforeAnimation);
    await d('auto-skeleton','empty').check();
    check('skeleton: empty content removes placeholders without disabling configuration', await d('auto-skeleton','preview').locator('[data-auto-skeleton-shape]').count() === 0 && await d('auto-skeleton','loading').isVisible());
    await d('auto-skeleton','empty').uncheck(); await d('auto-skeleton','loading').uncheck();
    await open('form');
    await setDiagnostics(true);
    await d('form','inspect-slots').click();
    check('form: label/hint refs', /Label: SPAN; hint: DIV; native target: \S+/.test(await d('form','slot-observation').innerText()));
    const label = d('form','slot-label'); const target = await label.evaluate(node => node.closest('label').htmlFor);
    check('form: native Label ID association', await page.locator(`[id="${target}"]`).count() === 1);
    await label.click();
    check('form: label click callback and native focus', (await d('form','slot-observation').innerText()).includes('label clicks: 1') && await page.locator(`[id="${target}"]`).evaluate(node => node === document.activeElement));
    check('form: callback hint update', (await d('form','slot-hint').innerText()).includes('label clicked 1 times'));
    const hintId = await d('form','slot-hint').evaluate(node => node.parentElement.id);
    check('form: hint describedby association', (await page.locator(`[id="${target}"]`).getAttribute('aria-describedby') || '').split(' ').includes(hintId));
    check('form: object fieldset legend title', await d('form','slot-section-title').evaluate(node => !!node.closest('legend')?.closest('fieldset')));
    await page.locator(`[id="${target}"]`).fill('Dirty slot notes'); await d('form','dirty-refresh').click();
    check('form: refresh retains dirty slot-labelled value', await page.locator(`[id="${target}"]`).inputValue() === 'Dirty slot notes');
    await d('form','new-record').click();
    check('form: new record resets slot-labelled value', await page.locator(`[id="${target}"]`).inputValue() === 'Editable notes');
    await page.locator(`[id="${target}"]`).fill('Reset slot notes'); await d('form','reset').click();
    check('form: reset restores slot-labelled field', await page.locator(`[id="${target}"]`).inputValue() === 'Editable notes');
    await d('form','submit').click();
    await page.waitForFunction(() => document.querySelector('[data-form-demo="save-count"]').textContent.includes('1'));
    check('form: submit callback remains connected', (await d('form','save-count').innerText()).includes('1'));
    await page.waitForFunction(() => document.querySelector('[data-form-demo="status"]').textContent.includes('pending: false'));
    check('form: delayed save completes', (await d('form','saved').innerText()).includes('Order B'));
    await page.screenshot({path:path.join(evidence,`migrated-${mode}-form.png`),fullPage:true});
    await page.locator('h1').scrollIntoViewIfNeeded();
    await page.screenshot({path:path.join(evidence,`migrated-${mode}-form-viewport.png`),fullPage:false});
    check('known Fluent ColorArea Chrome warning occurs once', peerWarnings.length === 1);
    fs.writeFileSync(path.join(evidence,`migrated-${mode}-result.json`), JSON.stringify({mode,browser:browser.version(),base,viewports:[[1280,900],[390,844]],checks,console:errors,peerWarnings},null,2));
    assert.deepEqual(errors, [], `${mode}: console errors/warnings`);
    console.log(JSON.stringify({mode,checks:checks.length,console:errors},null,2));
  } finally { await browser.close(); server.close(); }
})().catch(error => { console.error(error); process.exitCode=1; });
