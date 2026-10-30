const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');
const { loadModule, cleanup } = require('./fixtures/formatted-text-harness.cjs');

const source = 'packages/spfx-react-toolkit/src/controls/auto-skeleton/auto-skeleton-geometry.internal.ts';
let geometry;

test.before(async () => {
  assert.ok(fs.existsSync(path.resolve(__dirname, '..', source)), 'The private AutoSkeleton scanner is missing');
  geometry = await loadModule(source);
});
test.after(cleanup);

function rect(left, top, width, height) {
  return { x: left, y: top, left, top, width, height, right: left + width, bottom: top + height, toJSON() { return {}; } };
}

function fixture(t, html = '', css = '') {
  const dom = new JSDOM('<!doctype html><html><head></head><body><div id="root"><div id="content"></div></div></body></html>');
  t.after(() => dom.window.close());
  const document = dom.window.document;
  const root = document.getElementById('root');
  const content = document.getElementById('content');
  const style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);
  content.innerHTML = html;
  root.getBoundingClientRect = () => rect(100, 50, 400, 200);
  const textRects = new Map();
  // Installed JSDOM has no Range layout implementation; double that geometry
  // boundary while exercising real DOM traversal, selectors and computed CSS.
  document.createRange = () => {
    let selected;
    return {
      selectNodeContents(node) { selected = node; },
      getClientRects() { return textRects.get(selected) || []; }
    };
  };
  function box(selector, bounds) {
    const element = content.querySelector(selector);
    assert.ok(element, `Missing test element ${selector}`);
    element.getBoundingClientRect = () => rect(...bounds);
    return element;
  }
  function text(selector, index, bounds) {
    const parent = selector ? content.querySelector(selector) : content;
    const node = [...parent.childNodes].filter(child => child.nodeType === 3)[index];
    assert.ok(node, `Missing text node ${selector || 'content'} ${index}`);
    textRects.set(node, bounds.map(values => rect(...values)));
  }
  return {
    document, root, content, box, text,
    scan(options = {}) {
      return geometry.scanAutoSkeleton(content, root, { minSize: 8, ignoreSelectors: [], ...options });
    }
  };
}

test('measures HTML visual leaves as rectangles and rounded circles', t => {
  const f = fixture(t, '<img id="wide"><img id="round"><img id="equal">', '#round { border-radius: 50%; }');
  f.box('#wide', [120, 70, 80, 30]);
  f.box('#round', [220, 80, 40, 40]);
  f.box('#equal', [300, 90, 24, 24]);
  assert.deepEqual(f.scan(), [
    { x: 20, y: 20, width: 80, height: 30, radius: 0, shape: 'rectangle' },
    { x: 120, y: 30, width: 40, height: 40, radius: 20, shape: 'circle' },
    { x: 200, y: 40, width: 24, height: 24, radius: 0, shape: 'rectangle' }
  ]);
});

test('uses actual text line rectangles and measures nested inline text once', t => {
  const f = fixture(t, '<p id="text">first <strong id="inline">second</strong> third</p>');
  f.text('#text', 0, [[110, 60, 60, 14], [110, 80, 35, 14]]);
  f.text('#inline', 0, [[175, 60, 45, 14]]);
  f.text('#text', 1, [[225, 60, 32, 14]]);
  assert.deepEqual(f.scan(), [
    { x: 12, y: 12, width: 56, height: 10, radius: 0, shape: 'rectangle' },
    { x: 12, y: 32, width: 31, height: 10, radius: 0, shape: 'rectangle' },
    { x: 77, y: 12, width: 41, height: 10, radius: 0, shape: 'rectangle' },
    { x: 127, y: 12, width: 28, height: 10, radius: 0, shape: 'rectangle' }
  ]);
});

test('measures direct content text without generating a wrapper bone', t => {
  const f = fixture(t, ' direct text ');
  f.text(undefined, 0, [[110, 60, 80, 16]]);
  assert.deepEqual(f.scan(), [{ x: 12, y: 12, width: 76, height: 12, radius: 0, shape: 'rectangle' }]);
});

test('separates touching text fragments without collapsing small eligible fragments', t => {
  const f = fixture(t, 'text');
  f.text(undefined, 0, [[110, 60, 20, 14], [130, 60, 20, 14], [110, 74, 20, 14], [150, 60, 2, 2]]);
  const bones = f.scan({ minSize: 1 });
  assert.equal(bones[1].x - (bones[0].x + bones[0].width), 4);
  assert.equal(bones[2].y - (bones[0].y + bones[0].height), 4);
  assert.deepEqual(bones[3], { x: 50, y: 10, width: 2, height: 2, radius: 0, shape: 'rectangle' });
  f.text(undefined, 0, [[110, 60, 8, 8]]);
  assert.deepEqual(f.scan(), [{ x: 12, y: 12, width: 4, height: 4, radius: 0, shape: 'rectangle' }]);
});

test('prunes ignored subtrees and selector matches before reading their geometry', t => {
  const f = fixture(t, '<div data-skeleton-ignore><img id="marker"></div><div class="skip"><img id="selector"></div><img id="keep">');
  f.box('#marker', [120, 60, 20, 20]).getBoundingClientRect = () => { throw new Error('Ignored subtree measured'); };
  f.box('#selector', [150, 60, 20, 20]).getBoundingClientRect = () => { throw new Error('Ignored subtree measured'); };
  f.box('#keep', [180, 60, 20, 20]);
  assert.deepEqual(f.scan({ ignoreSelectors: ['.skip'] }), [{ x: 80, y: 10, width: 20, height: 20, radius: 0, shape: 'rectangle' }]);
});

test('honors authored hidden ancestors, including opacity, visibility and display', t => {
  const f = fixture(t,
    '<div class="transparent"><img id="opacity"></div><div class="invisible"><img id="visibility" class="visible"></div><div class="gone"><img id="display"></div><img id="keep">',
    '.transparent { opacity: 0; } .invisible { visibility: hidden; } .visible { visibility: visible; } .gone { display: none; }');
  for (const selector of ['#opacity', '#visibility', '#display']) {
    f.box(selector, [120, 60, 20, 20]).getBoundingClientRect = () => { throw new Error('Hidden subtree measured'); };
  }
  f.box('#keep', [180, 60, 20, 20]);
  assert.deepEqual(f.scan(), [{ x: 80, y: 10, width: 20, height: 20, radius: 0, shape: 'rectangle' }]);
});

test('ignores owned wrapper opacity while preserving authored root and wrapper hiding', t => {
  const f = fixture(t, '<img id="leaf">', '#content { opacity: 0; } .gone { display: none; } .invisible { visibility: hidden; } .transparent { opacity: 0; }');
  f.box('#leaf', [110, 60, 30, 20]);
  assert.equal(f.scan().length, 1);
  for (const className of ['gone', 'invisible', 'transparent']) {
    f.root.className = className;
    assert.deepEqual(f.scan(), []);
  }
  f.root.className = '';
  f.content.className = 'gone';
  assert.deepEqual(f.scan(), []);
});

test('honors hiding outside the root as well as inside it', t => {
  const f = fixture(t, '<img id="leaf">', 'body { opacity: 0; }');
  f.box('#leaf', [110, 60, 30, 20]);
  assert.deepEqual(f.scan(), []);
});

test('forced shapes are terminal and marked containers measure only their content', t => {
  const f = fixture(t, '<div id="forced" data-skeleton-shape="rectangle"><img id="inside"></div><div id="circle" data-skeleton-shape="circle"><span>text</span></div><div id="container" data-skeleton-container><img id="child"></div>');
  f.box('#forced', [110, 60, 90, 30]);
  f.box('#inside', [115, 65, 40, 20]).getBoundingClientRect = () => { throw new Error('Forced child measured'); };
  f.box('#circle', [220, 60, 40, 40]);
  f.box('#container', [300, 60, 90, 80]);
  f.box('#child', [305, 65, 25, 15]);
  assert.deepEqual(f.scan(), [
    { x: 10, y: 10, width: 90, height: 30, radius: 0, shape: 'rectangle' },
    { x: 120, y: 10, width: 40, height: 40, radius: 20, shape: 'circle' },
    { x: 205, y: 15, width: 25, height: 15, radius: 0, shape: 'rectangle' }
  ]);
});

test('unsupported shape markers do not force a container or expose a third shape', t => {
  const f = fixture(t, '<div id="legacy" data-skeleton-shape="rect"><img id="child"></div><div data-skeleton-shape="square"><img id="equal"></div>');
  f.box('#legacy', [110, 60, 90, 30]);
  f.box('#child', [120, 70, 30, 15]);
  f.box('#equal', [160, 70, 24, 24]);
  assert.deepEqual(f.scan(), [
    { x: 20, y: 20, width: 30, height: 15, radius: 0, shape: 'rectangle' },
    { x: 60, y: 20, width: 24, height: 24, radius: 0, shape: 'rectangle' }
  ]);
});

test('normalizes pixels before minimum-size filtering and deduplicates identical geometry', t => {
  const f = fixture(t, '<img id="small"><img id="threshold"><img id="same">');
  f.box('#small', [110.2, 60.1, 7.4, 20]);
  f.box('#threshold', [120.4, 70.4, 7.6, 8.4]);
  f.box('#same', [120.49, 70.49, 8, 8]);
  assert.deepEqual(f.scan(), [{ x: 20, y: 20, width: 8, height: 8, radius: 0, shape: 'rectangle' }]);
  assert.deepEqual(f.scan({ minSize: 9 }), []);
});

test('subtracts root borders and adds root scroll for overlay-local coordinates', t => {
  const f = fixture(t, '<img id="leaf">');
  Object.defineProperties(f.root, { clientLeft: { value: 3 }, clientTop: { value: 5 }, scrollLeft: { value: 11 }, scrollTop: { value: 17 } });
  f.box('#leaf', [120, 80, 40, 20]);
  assert.deepEqual(f.scan(), [{ x: 28, y: 42, width: 40, height: 20, radius: 0, shape: 'rectangle' }]);
});

test('rejects nonfinite and zero measured geometry without poisoning other bones', t => {
  const f = fixture(t, '<img id="bad"><img id="zero"><img id="keep">');
  f.box('#bad', [NaN, 60, Infinity, 20]);
  f.box('#zero', [120, 60, 0, 20]);
  f.box('#keep', [180, 60, 20, 20]);
  assert.deepEqual(f.scan(), [{ x: 80, y: 10, width: 20, height: 20, radius: 0, shape: 'rectangle' }]);
  f.root.getBoundingClientRect = () => rect(NaN, 50, 400, 200);
  assert.deepEqual(f.scan(), []);
});

test('rejects coordinate arithmetic that overflows despite finite input rectangles', t => {
  const f = fixture(t, '<img id="leaf">');
  f.root.getBoundingClientRect = () => rect(-1e308, 50, 400, 200);
  f.box('#leaf', [1e308, 60, 30, 20]);
  assert.deepEqual(f.scan(), []);
});

test('empty or entirely hidden content does not invent a wrapper or parent bone', t => {
  const f = fixture(t, '<div><img class="gone"></div> \n ', '.gone { display: none; }');
  f.content.getBoundingClientRect = () => rect(100, 50, 400, 200);
  assert.deepEqual(f.scan(), []);
  f.content.textContent = '';
  assert.deepEqual(f.scan(), []);
});

test('uses the owner document and leaves the DOM unchanged', async t => {
  const f = fixture(t, '<img id="leaf">');
  f.box('#leaf', [110, 60, 30, 20]);
  const mutations = [];
  const observer = new f.document.defaultView.MutationObserver(records => mutations.push(...records));
  observer.observe(f.root, { attributes: true, childList: true, characterData: true, subtree: true });
  t.after(() => observer.disconnect());
  assert.deepEqual(f.scan(), [{ x: 10, y: 10, width: 30, height: 20, radius: 0, shape: 'rectangle' }]);
  await Promise.resolve();
  assert.equal(mutations.length, 0);
});

test('validates positive finite minSize and CSS selector syntax explicitly', t => {
  const f = fixture(t);
  for (const value of [0, -1, NaN, Infinity, -Infinity]) {
    assert.throws(() => geometry.validateAutoSkeletonOptions(f.content, value, []), RangeError);
  }
  assert.throws(() => geometry.validateAutoSkeletonOptions(f.content, 8, ['[broken']), /selector/i);
  assert.doesNotThrow(() => geometry.validateAutoSkeletonOptions(f.content, 0.5, ['.absent', '[data-skeleton-ignore]', ':scope > p']));
});
