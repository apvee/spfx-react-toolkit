const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const read = file => fs.readFileSync(path.resolve(__dirname, '..', file), 'utf8');
const base = 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/';

test('groups have lazy interactive coverage without duplicate hook/provider inventory', () => {
  const registry = read(base + 'demoRegistry.ts');
  const panel = read(base + 'panels/GroupsPanel.tsx');
  assert.match(registry, /key: 'groups'/);
  assert.match(registry, /webpackChunkName: 'spfx-demo-groups'/);
  assert.match(registry, /import\(.*GroupsPanel/);
  for (const symbol of ['SettingsGroup', 'AccordionGroup']) {
    assert.equal([...registry.matchAll(new RegExp(`symbol: '${symbol}'`, 'g'))].length, 1);
    assert.ok(panel.includes(`${symbol}.Item`));
  }
  for (const marker of ['theme', 'direction', 'resize', 'multiple', 'unmount', 'hidden', 'disabled', 'settings', 'accordion', 'email', 'note', 'close-all', 'inspect', 'observations']) {
    assert.ok(panel.includes(`data-groups-demo="${marker}"`), marker);
  }
  for (const marker of ['Input aria-label=', 'Switch aria-label=', 'Icon iconName=', 'onToggle=', 'openItems=', 'setOpenItems([])', 'unmountOnClose=', 'defaultValue="Retained note"', 'title={<span>', 'action=', 'ref=']) {
    assert.ok(panel.includes(marker), marker);
  }
  assert.match(panel, /\[320, 500, 639, 640, 900\]/);
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
  assert.doesNotMatch(panel, /style=|Layout\.|\/lib\//);
});

test('group guides explain labels, trigger restrictions and separate host execution', () => {
  const settings = read('docs/api/controls/settings-group.md');
  const accordion = read('docs/api/controls/accordion-group.md');
  assert.match(settings, /not implicit labels/);
  assert.match(settings, /wrap/);
  assert.match(accordion, /Never place interactive JSX/);
  assert.match(accordion, /unmountOnClose/);
  assert.match(accordion, /remain mounted/);
  assert.match(accordion, /native inert/);
  assert.match(read('docs/SHAREPOINT-VALIDATION.md'), /## SettingsGroup and AccordionGroup/);
});
