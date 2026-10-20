const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');

test('AutoSkeleton has a lazy demo with observable props, geometry and retained content', () => {
  const base = 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/';
  const panel = read(`${base}panels/AutoSkeletonPanel.tsx`);
  const registry = read(`${base}demoRegistry.ts`);
  assert.match(registry, /key: 'auto-skeleton'/);
  assert.match(registry, /symbol: 'AutoSkeleton'/);
  assert.match(registry, /import\(.*AutoSkeletonPanel/);
  for (const marker of ['loading', 'animation', 'minSize', 'ignoreSelectors', 'resize', 'empty',
    'direction', 'preview', 'inspect-ref', 'observations', 'focus-load', 'focus-target', 'circle', 'rectangle', 'mixed-text']) {
    assert.ok((panel.includes(`data-auto-skeleton-demo="${marker}"`) || panel.includes(`'data-auto-skeleton-demo': '${marker}'`)), marker);
  }
  assert.match(panel, /data-skeleton-ignore/);
  assert.match(panel, /data-skeleton-shape="circle"/);
  assert.match(panel, /data-skeleton-shape="rectangle"/);
  assert.match(panel, /width: '160px', height: '24px'/);
  assert.match(panel, /width: '40px', height: '40px'/);
  assert.match(panel, /useStableCallback/);
  assert.match(panel, /ref=\{rootRef\}/);
  assert.match(panel, /onClick=/);
  assert.doesNotMatch(panel, /style=|shape="square"|sessionStorage|useRef.*null/);
});

test('source and compiled sample browser fixtures include AutoSkeleton', () => {
  assert.match(read('tests/fixtures/sx-browser/index.tsx'), /<AutoSkeletonFixture/);
  assert.match(read('tests/fixtures/styles-panel-browser/index.tsx'), /<AutoSkeletonPanel/);
  assert.match(read('tests/fixtures/styles-panel-browser/toolkit-boundary.ts'), /lib\/controls/);
  const fixture = read('tests/fixtures/sx-browser/auto-skeleton.tsx');
  assert.match(fixture, /window\.autoSkeletonFixture/);
  assert.match(fixture, /focusAndLoad/);
  assert.match(fixture, /data-skeleton-shape="rectangle"/);
  assert.match(fixture, /data-skeleton-shape="circle"/);
  assert.doesNotMatch(fixture, /style=|shape="square"/);
});

// Configuration controls belong to Fluent and remain separate from the demonstrated control.
test('AutoSkeleton configuration uses Fluent native slots and retained diagnostic regions', () => {
  const panel = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/AutoSkeletonPanel.tsx'), 'utf8');
  assert.doesNotMatch(panel, /<(?:input|select|textarea|button|details)\b|style=|from '@fluentui\/react'/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /defaultOpenItems/);
  assert.match(panel, /input=\{\{/);
  assert.match(panel, /onChange=\{\(_event, data\)/);
});
