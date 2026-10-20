// Product acceptance uses the compiled actual sample panels. Legacy comparisons
// remain on the normal full fixture route and retain their existing assertions.
async function checkFoundationSingleBrowser(page, check) {
  const origin = new URL(page.url()).origin;
  const settle = () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  const route = async name => {
    await page.goto(`${origin}/?foundation-product=${name}`);
    await page.locator('[data-demo-section]').waitFor();
    check(`${name}: one configuration`, await page.locator('[data-demo-configuration]').count(), 1);
    check(`${name}: one section`, await page.locator('[data-demo-section]').count(), 1);
  };
  const open = async () => { await page.getByRole('button', { name: 'Advanced diagnostics' }).click(); };
  const choice = name => page.locator(`select[data-demo-variant="${name}"]`);
  await route('styles');
  const sx = id => page.locator(`[data-sx-demo="${id}"]`);
  await sx('layout-width').fill('310');
  const retained = await sx('layout-preview').getAttribute('class');
  for (const [variant, subject] of [['roles','role-preview'],['preset','preset-preview'],['catalog','catalog-preview'],['interaction','state-preview'],['container','query-a-preview'],['viewport','viewport-preview'],['scrollbar','scroll-preview']]) {
    await choice('styles').selectOption(variant);
    check(`styles ${variant}: actual selected subject`, await sx(subject).count(), 1);
    check(`styles ${variant}: one configuration`, await page.locator('[data-demo-configuration]').count(), 1);
    check(`styles ${variant}: no comparison container`, await sx('query-b').count(), 0);
  }
  await choice('styles').selectOption('catalog');
  await sx('catalog-family').selectOption('width'); await sx('catalog-member').selectOption('px');
  await sx('catalog-value').fill('84');
  await sx('catalog-family').selectOption('foreground'); await sx('catalog-member').selectOption('primary');
  check('Styles constant descriptor has no numeric field', await sx('catalog-value').count(), 0);
  await sx('catalog-family').selectOption('width'); await sx('catalog-member').selectOption('px');
  check('Styles factory value retained', await sx('catalog-value').inputValue(), '84');
  await choice('styles').selectOption('interaction');
  await sx('selected').check();
  check('styles selected: actual aria pressed', await sx('state-preview').getAttribute('aria-pressed'), 'true');
  await sx('disabled').check();
  check('styles disabled: actual button', await sx('state-preview').isDisabled(), true);
  await page.evaluate(() => window.demoFixtureHost.update({ theme: 'contrast' }));
  await choice('styles').selectOption('preset'); await sx('preset').selectOption('inverted');
  check('Styles current inverted preset shows high contrast fallback', (await sx('preset-notice').textContent()).includes('Canvas is shown instead'), true);
  for (const scenario of ['layout','roles']) {
    await choice('styles').selectOption(scenario);
    check(`Styles ${scenario}: inactive preset notice absent`, await sx('preset-notice').count(), 0);
  }
  await choice('styles').selectOption('preset');
  check('Styles preset retained after other scenarios', await sx('preset').inputValue(), 'inverted');
  await page.evaluate(() => window.demoFixtureHost.update({ theme: 'light' }));
  await choice('styles').selectOption('layout');
  check('styles retains width draft after scenario changes', await sx('layout-width').inputValue(), '310');
  check('styles retains applied descriptor', await sx('layout-preview').getAttribute('class'), retained);
  await route('formatted-text');
  const ft = id => page.locator(`[data-formatted-text-demo="${id}"]`);
  await ft('text').fill('Il **progetto** è pronto per la revisione.');
  check('FormattedText actual Markdown', await ft('preview').locator('strong').textContent(), 'progetto');
  await open();
  check('FormattedText diagnostics contain no second subject', await ft('preview').count(), 1);
  await choice('formatted-text').selectOption('jsx');
  check('FormattedText JSX is in current subject', await page.locator('[data-demo-subject] [data-formatted-text-demo="mixed-jsx"] button').count(), 1);
  await page.locator('[data-demo-subject] [data-formatted-text-demo="mixed-jsx"] button').click();
  check('FormattedText one JSX activation counted once', (await ft('observations').textContent()).includes('Preview clicks: 1.'), true);
  await page.locator('[data-demo-subject] [data-formatted-text-demo="mixed-jsx"] button').press('Enter');
  check('FormattedText keyboard activation counted once', (await ft('observations').textContent()).includes('Preview clicks: 2.'), true);
  await choice('formatted-text').selectOption('text');
  check('FormattedText returns to retained text', await ft('preview').locator('strong').textContent(), 'progetto');
  await route('relative-time');
  const rt = id => page.locator(`[data-relative-time-demo="${id}"]`);
  await open();
  check('RelativeTime one actual time with diagnostics open', await page.locator('[data-demo-subject] time').count(), 1);
  await choice('relative-time').selectOption('invalid');
  check('RelativeTime invalid has no time node', await page.locator('[data-demo-subject] time').count(), 0);
  check('RelativeTime actual fallback', await page.locator('[data-demo-subject]').textContent(), 'Date unavailable');
  await choice('relative-time').selectOption('slot');
  check('RelativeTime one actual slot fallback', await rt('fallback-slot').count(), 1);
  await rt('fallback-slot').click();
  await rt('inspect-slot').click();
  check('RelativeTime actual fallback ref and click', (await page.locator('[data-demo-configuration]').textContent()).includes('Fallback element: span. Fallback clicks: 1.'), true);
  await choice('relative-time').selectOption('event');
  check('RelativeTime event restored', await page.locator('[data-demo-subject] time').count(), 1);
  await route('auto-skeleton');
  const auto = id => page.locator(`[data-auto-skeleton-demo="${id}"]`);
  await auto('focus-target').fill('Note conservata durante il caricamento');
  await auto('loading').check();
  check('AutoSkeleton busy state on actual root', await auto('preview').getAttribute('aria-busy'), 'true');
  await auto('animation').selectOption('none');
  await settle();
  check('AutoSkeleton static Fluent items on actual root', await auto('preview').locator('.fui-SkeletonItem').evaluateAll(items => items.length > 0 && items.every(item => {
    const css = getComputedStyle(item, '::after');
    return css.animationName === 'none' && css.backgroundImage === 'none' && css.transform === 'none';
  })), true);
  check('AutoSkeleton static busy state retained', await auto('preview').getAttribute('aria-busy'), 'true');
  for (const animation of ['wave', 'pulse']) {
    await auto('animation').selectOption(animation);
    await settle();
    check(`AutoSkeleton ${animation} restored after static mode`, await auto('preview').locator('.fui-SkeletonItem').first().evaluate(item => getComputedStyle(item, '::after').animationName !== 'none'), true);
  }
  await auto('loading').uncheck();
  check('AutoSkeleton child note retained', await auto('focus-target').inputValue(), 'Note conservata durante il caricamento');
  await open();
  check('AutoSkeleton one actual subject with diagnostics open', await auto('preview').count(), 1);
  await route('state-message');
  const state = id => page.locator(`[data-state-message-demo="${id}"]`);
  await state('title').fill('Updated available');
  check('StateMessage editable actual heading', await state('preview').locator('h3').textContent(), 'Updated available');
  await choice('state-message').selectOption('success');
  check('StateMessage subject remains single after preset change', await state('preview').count(), 1);
  await open();
  await state('slots').check();
  check('StateMessage actual slot heading', await state('preview').locator('[data-state-message-demo="title-slot"]').textContent(), '**Literal slot title**');
  await route('imports');
  for (const adapter of ['root','domain','legacy','root']) {
    await choice('imports').selectOption(adapter);
    const result = page.locator(`[data-imports-demo="${adapter}"]`);
    for (const other of ['root','domain','legacy'].filter(item => item !== adapter)) check(`Imports ${adapter}: ${other} absent`, await page.locator(`[data-imports-demo="${other}"]`).count(), 0);
    await result.getByRole('button', { name: 'Increment counter', exact: true }).click();
    await result.getByRole('button', { name: 'Increment counter', exact: true }).click();
    await result.getByRole('button', { name: 'Invoke captured callback', exact: true }).click();
    check(`Imports ${adapter}: captured callback latest value`, (await result.textContent()).includes('Observed counter: 2'), true);
    check(`Imports ${adapter}: captured callback stable identity`, (await result.textContent()).includes('Same callback identity: yes'), true);
    check(`Imports ${adapter}: one result inside configuration`, await page.locator('[data-demo-result]').count(), 1);
  }
}
module.exports = { checkFoundationSingleBrowser };
