const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const crypto = require('node:crypto');
const harness = require('./fixtures/formatted-text-harness.cjs');
const entry = path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/state-message/illustrations/illustration-registry.internal.ts');
// Fingerprints of the original OfficeDev geometry and palette, excluding instance-specific paint IDs.
const originalArtwork = {
  "generic": {
    "light": {
      "viewBox": "0 0 280 173.2",
      "paths": "f9b7b48e94b0b16fa284f92c5fb2de4c92a5a37a8d8a9b4261d552157d65fdb8",
      "appearance": "3355cc070d6c445a81e01171d55c5baf6005fddba1ef94f37d840cf66761a676"
    },
    "dark": {
      "viewBox": "0 0 280 173.2",
      "paths": "beb5e760e6caf2600d31a4280ef3daabbc6eb81e72fbfd437b3a485776a35417",
      "appearance": "88bff0a5c5e2f4de1da6869530d7473160c4e064d96ca2679bf3846acb673e8f"
    }
  },
  "empty": {
    "light": {
      "viewBox": "0 0 252 182.2",
      "paths": "f72ae288b0570f2f4b3a38598422da8dd1134997f5a6593a4faf2af4af64e0e1",
      "appearance": "f78d3b14bc87fc112ce8e6c0e2ea819a6dc46bca3130c1b0a73041a4549aec56"
    },
    "dark": {
      "viewBox": "0 0 252 182.2",
      "paths": "205dbc551547f8295e76307dff8aa19e3f83232d56e819c79c4114eae98f9bcc",
      "appearance": "53e1938ebc29b89e5fcb68cd584c0dac6166739dd0f1c0b767d80933259bd259"
    }
  },
  "error": {
    "light": {
      "viewBox": "0 0 121.6 122.2",
      "paths": "5634b90c684a0cb6489dda7ff4449a28b41ba6df038162e6ad3270843950a33a",
      "appearance": "31c67ef7a3938f80aab032eb6c2adbb21836e8fa5e1dc84951ea7a862f9d7c8e"
    },
    "dark": {
      "viewBox": "0 0 121.6 122.2",
      "paths": "5634b90c684a0cb6489dda7ff4449a28b41ba6df038162e6ad3270843950a33a",
      "appearance": "740bbbd49e81b9bf04c4b1b41df607dc0b743dfaf007c8e2cff18761f459a2cc"
    }
  },
  "success": {
    "light": {
      "viewBox": "0 0 119.8 120",
      "paths": "670063e1107453151396e8a1791ed36fcb9526b7ad7017af3e883ab9fd57bbc1",
      "appearance": "a37dc12e34afdc34d3efee54e529b383d9e4acced8ec6fabd43c4ca565953930"
    },
    "dark": {
      "viewBox": "0 0 117.7 120",
      "paths": "0e956190194f6332ce2f014242bef185b02e8d1af1765d895863529b8b692f4d",
      "appearance": "77642b424c3828f318b8572f0c1d241d8780a2f21eecefe4059489611e49f0a5"
    }
  },
  "welcome": {
    "light": {
      "viewBox": "0 0 193 219",
      "paths": "9022a4b42fea41edc1af23aacb22e8d3dd789dada745da553f2738985c6c066c",
      "appearance": "ffc4791fb3b759f420ac5c611e9c1affda5200e9438daf20588b3a61b7214889"
    },
    "dark": {
      "viewBox": "0 0 193 219",
      "paths": "b359c51af6929c0eba05a383031d3e86f4fd156b17fb976c24775e8c48a60b93",
      "appearance": "f496844afb8a87155049fc3b7e3878a15a9423aea4d5f964ebaea98888600881"
    }
  }
};
let React, ReactDOM, act, resolvePresetIllustration, dom, RendererProvider, renderer;
test.before(async () => {
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ RendererProvider } = require('@griffel/react'));
  // JSDOM cannot parse @container; real geometry is checked by the browser runner.
  renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  ({ resolvePresetIllustration } = await harness.loadModule(entry));
});
test.after(async () => { await harness.cleanup(); dom?.window.close(); });
test('all presets preserve both original geometries and decorative SVG semantics', () => {
  const host = document.createElement('div'); document.body.appendChild(host);
  try {
    for (const [preset, modes] of Object.entries(originalArtwork)) {
      const Illustration = resolvePresetIllustration(preset);
      for (const [mode, expected] of Object.entries(modes)) {
        act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(Illustration, { mode })), host); });
        const svg = host.querySelector('svg'); assert.ok(svg);
        assert.equal(svg.getAttribute('viewBox'), expected.viewBox);
        assert.equal(svg.getAttribute('preserveAspectRatio'), 'xMidYMid meet');
        assert.equal(svg.getAttribute('aria-hidden'), 'true');
        assert.equal(svg.getAttribute('focusable'), 'false');
        assert.equal(svg.querySelector('title'), null);
        assert.equal(svg.querySelector('[style]'), null);
        const paths = Array.from(svg.querySelectorAll('path'), element => element.getAttribute('d')).join('\n');
        assert.equal(crypto.createHash('sha256').update(paths).digest('hex'), expected.paths);
        const attributes = ['d', 'points', 'transform', 'cx', 'cy', 'r', 'rx', 'ry', 'x', 'y', 'x1', 'y1', 'x2', 'y2', 'width', 'height', 'fill', 'stroke', 'stopColor', 'offset', 'opacity', 'fillRule', 'strokeWidth'];
        const appearance = [];
        for (const element of [svg, ...svg.querySelectorAll('*')]) for (const attribute of attributes) {
          const svgAttribute = attribute.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`);
          const value = element.getAttribute(svgAttribute);
          if (value !== null && !value.startsWith('url(')) appearance.push(`${attribute}:${value}`);
        }
        assert.equal(crypto.createHash('sha256').update(appearance.sort().join('\n')).digest('hex'), expected.appearance);
      }
    }
  } finally { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); }
});
test('repeated illustrations have unique IDs and paint references in their own SVG', () => {
  const host = document.createElement('div'); document.body.appendChild(host);
  try {
    const children = [];
    for (const preset of Object.keys(originalArtwork)) for (const mode of ['light', 'dark']) for (let instance = 0; instance < 2; instance++) {
      children.push(React.createElement(resolvePresetIllustration(preset), { mode, key: `${preset}-${mode}-${instance}` }));
    }
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement('div', undefined, children)), host); });
    const documentIds = new Set(); let references = 0;
    for (const svg of host.querySelectorAll('svg')) {
      const localIds = new Set(Array.from(svg.querySelectorAll('[id]'), element => element.id));
      for (const id of localIds) { assert.equal(documentIds.has(id), false, `duplicate ID ${id}`); documentIds.add(id); }
      for (const [, reference] of svg.outerHTML.matchAll(/url\(#([^)]+)\)/g)) {
        assert.ok(localIds.has(reference), `paint reference ${reference} is outside this SVG`); references++;
      }
    }
    assert.ok(references > 0);
  } finally { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); }
});
