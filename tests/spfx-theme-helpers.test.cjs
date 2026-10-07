const test = require('node:test');
const assert = require('node:assert/strict');
const { createTheme } = require('@fluentui/react');
const { createV9Theme } = require('@fluentui/react-migration-v8-v9');
const { webLightTheme, teamsLightTheme, teamsDarkTheme, teamsHighContrastTheme } = require('@fluentui/react-theme');
const { loadSxModules } = require('./fixtures/sx-harness.cjs');
const { createFluent9ThemeFromSPFxTheme, getTeamsFluentTheme } = loadSxModules().load('../spfx-theme.helpers');

test('SPFx conversion restores host neutral hover and pressed colors after the real shim collapses them', () => {
  const host = createTheme();
  const upstream = createV9Theme(host);
  assert.equal(upstream.colorNeutralStrokeAccessibleHover, '#605e5c');
  assert.equal(upstream.colorNeutralStrokeAccessiblePressed, '#605e5c');
  const converted = createFluent9ThemeFromSPFxTheme(host);
  assert.equal(converted.colorNeutralStrokeAccessible, '#605e5c');
  assert.equal(converted.colorNeutralStrokeAccessibleHover, '#323130');
  assert.equal(converted.colorNeutralStrokeAccessiblePressed, '#201f1e');
});

test('custom inverted host palettes preserve source data and every other converted token', () => {
  const host = createTheme({ isInverted: true, palette: {
    neutralSecondary: '#a1a2a3', neutralPrimary: '#b1b2b3', neutralDark: '#c1c2c3'
  } });
  const before = JSON.stringify(host);
  Object.freeze(host.palette);
  Object.freeze(host);
  const upstream = createV9Theme(host), converted = createFluent9ThemeFromSPFxTheme(host);
  assert.equal(converted.colorNeutralStrokeAccessible, '#a1a2a3');
  assert.equal(converted.colorNeutralStrokeAccessibleHover, '#b1b2b3');
  assert.equal(converted.colorNeutralStrokeAccessiblePressed, '#c1c2c3');
  for (const token of Object.keys(upstream)) {
    if (token !== 'colorNeutralStrokeAccessibleHover' && token !== 'colorNeutralStrokeAccessiblePressed') {
      assert.deepEqual(converted[token], upstream[token], `Unrelated converted token ${token} must stay unchanged`);
    }
  }
  assert.equal(JSON.stringify(host), before);
});

test('missing or empty host state colors preserve each upstream converted fallback', () => {
  for (const [colors, hover, pressed] of [
    [{ neutralPrimary: undefined, neutralDark: '' }, '#605e5c', '#605e5c'],
    [{ neutralPrimary: '   ', neutralDark: undefined }, '#605e5c', '#605e5c'],
    [{ neutralPrimary: '#123456', neutralDark: '' }, '#123456', '#605e5c'],
    [{ neutralPrimary: '', neutralDark: '#abcdef' }, '#605e5c', '#abcdef']
  ]) {
    const host = createTheme();
    host.palette = { ...host.palette, ...colors };
    const converted = createFluent9ThemeFromSPFxTheme(host);
    assert.equal(converted.colorNeutralStrokeAccessibleHover, hover);
    assert.equal(converted.colorNeutralStrokeAccessiblePressed, pressed);
  }
});

test('undefined SPFx theme and Teams themes keep their existing shared identities', () => {
  assert.equal(createFluent9ThemeFromSPFxTheme(undefined), webLightTheme);
  for (const [name, theme] of [[undefined, teamsLightTheme], ['default', teamsLightTheme],
    ['unknown', teamsLightTheme], ['DARK', teamsDarkTheme], ['contrast', teamsHighContrastTheme], ['highcontrast', teamsHighContrastTheme]]) {
    assert.equal(getTeamsFluentTheme(name), theme);
  }
});
