const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const harness = require('./fixtures/formatted-text-harness.cjs');
const entry = path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/state-message/state-message-theme.internal.ts');
let React, ReactDOM, act, resolveThemeMode, useStateMessageThemeMode, dom;
test.before(async () => {
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ resolveThemeMode, useStateMessageThemeMode } = await harness.loadModule(entry));
});
test.after(async () => { await harness.cleanup(); dom?.window.close(); });
test('classifies supported theme colors and falls back to light', () => {
  const { webLightTheme, webDarkTheme } = require('@fluentui/react-theme');
  assert.equal(resolveThemeMode(webLightTheme), 'light'); assert.equal(resolveThemeMode(webDarkTheme), 'dark');
  assert.equal(resolveThemeMode(undefined), 'light'); assert.equal(resolveThemeMode({}), 'light');
  for (const [background, foreground, expected] of [
    ['#fff', '#222', 'light'], ['#222', '#fff', 'dark'],
    ['#2228', '#fff8', 'dark'], ['#ffffff', '#242424', 'light'],
    ['#00000020', '#ffffffff', 'dark'], ['#777', '#777', 'light'],
    ['rgb(255, 255, 255)', 'rgb(36, 36, 36)', 'light'],
    ['rgba(41, 41, 41, 0.5)', 'rgba(255, 255, 255, 0.5)', 'dark'],
    ['rgb(0%, 0%, 0%)', 'rgba(100%, 100%, 100%, 50%)', 'dark'],
  ]) assert.equal(resolveThemeMode({ colorNeutralBackground1: background, colorNeutralForeground1: foreground }), expected);
  for (const invalid of ['transparent', 'Canvas', 'var(--background)', '#12', 'rgb(1, 2)', 'rgb(-1, 2, 3)', 'rgb(256, 2, 3)', 'rgba(1, 2, 3, 2)', 'rgb(1junk, 2, 3)']) {
    assert.equal(resolveThemeMode({ colorNeutralBackground1: invalid, colorNeutralForeground1: '#fff' }), 'light', invalid);
    assert.equal(resolveThemeMode({ colorNeutralBackground1: '#000', colorNeutralForeground1: invalid }), 'light', invalid);
  }
});
test('uses the nearest FluentProvider and updates theme without remounting content', () => {
  const { FluentProvider } = require('@fluentui/react-provider');
  const { webLightTheme, webDarkTheme } = require('@fluentui/react-theme');
  const host = document.createElement('div'); document.body.appendChild(host); let mounts = 0;
  function Probe() { React.useEffect(() => { mounts++; }, []); return React.createElement('output', undefined, useStateMessageThemeMode()); }
  function render(theme) {
    act(() => { ReactDOM.render(React.createElement(FluentProvider, { theme: webDarkTheme }, React.createElement(FluentProvider, { theme }, React.createElement(Probe))), host); });
  }
  try {
    act(() => { ReactDOM.render(React.createElement(Probe), host); }); assert.equal(host.textContent, 'light');
    render(webLightTheme); assert.equal(host.textContent, 'light'); const probe = host.querySelector('output'); const initialMounts = mounts;
    render(webDarkTheme); assert.equal(host.textContent, 'dark'); assert.equal(host.querySelector('output'), probe); assert.equal(mounts, initialMounts);
    render({ ...webLightTheme, colorNeutralBackground1: '#222', colorNeutralForeground1: '#fff' }); assert.equal(host.textContent, 'dark');
  } finally { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); }
});
