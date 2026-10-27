const assert = require('node:assert/strict');
const { test, before, after } = require('node:test');
const { loadModule, cleanup, createSxDom } = require('./fixtures/formatted-text-harness.cjs');
let content, dom, React, ReactDOM, act;
const options = { markdown: true, codeClassName: 'test-code', linkClassName: 'test-link' };
before(async () => {
  dom = createSxDom();
  React = require('react'); ReactDOM = require('react-dom');
  ({ act } = require('react-dom/test-utils'));
  content = await loadModule('packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text-content.internal.tsx');
});
after(async () => { await cleanup(); dom?.window.close(); });
function render(children, overrides) {
  const container = document.createElement('div'); document.body.appendChild(container);
  act(() => { ReactDOM.render(React.createElement('span', null,
    content.renderFormattedTextChildren(children, { ...options, ...overrides })), container); });
  return { container, close() { act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); } };
}
test('real Markdown renders inline grammar, whitespace, escape and hard breaks', () => {
  const view = render('  *em* **strong** ~~gone~~ `code`\\*literal\\*\nsoft  \nhard  ');
  try {
    assert.equal(view.container.querySelector('em').textContent, 'em');
    assert.equal(view.container.querySelector('strong').textContent, 'strong');
    assert.equal(view.container.querySelector('del').textContent, 'gone');
    assert.equal(view.container.querySelector('code').className, 'test-code');
    assert.equal(view.container.querySelectorAll('br').length, 1);
    assert.equal(view.container.textContent, '  em strong gone code*literal*\nsofthard  ');
    assert.equal(content.parseInlineMarkdown(' \t\n ').kind, 'parsed');
    const whitespace = render(' \t\n '); assert.equal(whitespace.container.textContent, ' \t\n '); whitespace.close();
  } finally { view.close(); }
});
test('unsupported syntax falls back to the entire exact source', () => {
  for (const source of ['# heading', 'first\n\nsecond', '- list', '- [x] task', '> quote', '---', '```js\nx\n```', '    code', '| a | b |\n| - | - |\n| 1 | 2 |', '![alt](image.png)', '[ref][id]\n\n[id]: https://example.com', '*good*\n\n- bad']) {
    const parsed = content.parseInlineMarkdown(source);
    assert.equal(parsed.kind, 'literal', source); assert.equal(parsed.source, source);
    const view = render(source); assert.equal(view.container.textContent, source); view.close();
  }
});
test('authored inline HTML remains literal React text', () => {
  const view = render('before <b onclick="alert(1)">raw</b> after');
  try { assert.equal(view.container.querySelector('b'), null); assert.equal(view.container.textContent, 'before <b onclick="alert(1)">raw</b> after'); }
  finally { view.close(); }
});
test('UTF-8, original AST node count and original depth are bounded', () => {
  assert.equal(content.parseInlineMarkdown('a'.repeat(65536)).kind, 'parsed');
  for (const source of ['a'.repeat(65537), 'é'.repeat(32769), '*a* '.repeat(700), '*a '.repeat(35) + 'x' + ' a*'.repeat(35)]) {
    const parsed = content.parseInlineMarkdown(source);
    assert.equal(parsed.kind, 'literal', source.slice(0, 70));
  }
  assert.equal(content.parseInlineMarkdown('é'.repeat(32768)).kind, 'parsed');
  assert.equal(content.parseInlineMarkdown('*a* '.repeat(666)).kind, 'parsed'); // exactly 2000 original nodes
  assert.equal(content.parseInlineMarkdown('*a* '.repeat(667)).reason, 'limit');
  assert.equal(content.parseInlineMarkdown('*'.repeat(58) + 'x' + '*'.repeat(58)).kind, 'parsed'); // depth 32
  assert.equal(content.parseInlineMarkdown('*'.repeat(60) + 'x' + '*'.repeat(60)).reason, 'limit');
});
test('URL policy rejects unsafe and malformed destinations', () => {
  for (const url of ['https://example.com/a', ' http://example.com ', 'mailto:a@example.com', 'tel:+39000', '/path', './path', '../path', 'path/to', '#hash', '?query=1']) assert.equal(content.isSafeMarkdownUrl(url), true, url);
  for (const url of [undefined, '', ' ', '//example.com', '\\path', 'https://', 'https://[', 'javascript:alert(1)', 'data:text/html,a', 'ftp://host', 'java\tscript:x', 'a b', 'a\u0000b', 'a\u007fb', 'https://example.com\\x']) assert.equal(content.isSafeMarkdownUrl(url), false, String(url));
});
test('safe links use real Fluent Link; decoded unsafe links retain descendants', () => {
  const view = render('[safe](https://example.com "Title") www.example.com [**bold**](javascript&#58;alert)');
  try {
    const links = view.container.querySelectorAll('a'); assert.equal(links.length, 2);
    assert.equal(links[0].getAttribute('href'), 'https://example.com');
    assert.equal(links[0].getAttribute('title'), 'Title'); assert.equal(links[0].getAttribute('target'), null);
    assert.ok(links[0].classList.contains('test-link')); assert.ok(links[0].classList.contains('fui-Link'));
    assert.equal(view.container.querySelector('strong').textContent, 'bold');
  } finally { view.close(); }
});
test('strings parse independently; only arrays and fragments are visited; markdown=false preserves identity', () => {
  const jsx = React.createElement('i', { title: 'opaque', key: 'jsx' }, '**untouched**');
  assert.equal(content.renderFormattedTextChildren(jsx, options), jsx);
  const children = ['*', jsx, '*', 3, null, React.createElement(React.Fragment, { key: 'fragment' }, '**yes**')];
  assert.equal(content.renderFormattedTextChildren(children, { ...options, markdown: false }), children);
  const view = render(children);
  try { assert.equal(view.container.querySelector('i').textContent, '**untouched**'); assert.equal(view.container.querySelector('em'), null); assert.equal(view.container.querySelector('strong').textContent, 'yes'); assert.ok(view.container.textContent.includes('3')); }
  finally { view.close(); }
});

test('a parser exception returns the complete source literal', async () => {
  const throwing = await loadModule('packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text-content.internal.tsx', { parserFailure: true });
  assert.deepEqual(throwing.parseInlineMarkdown('**original**'), { kind: 'literal', source: '**original**', reason: 'parse' });
});

test('static unkeyed JSX and numeric keyed siblings remain opaque without key warnings', () => {
  const messages = [];
  const originalError = console.error;
  console.error = (...args) => { messages.push(args.join(' ')); };
  let view;
  try {
    const button = React.createElement('button', { type: 'button' }, '**opaque**');
    const fragment = React.createElement(React.Fragment, null, '*parsed*');
    const keyed = React.createElement('i', { key: '1' }, 'keyed');
    const children = [keyed, '**first** ', button, fragment];
    view = render(children);
    assert.equal(view.container.querySelector('button').textContent, '**opaque**');
    assert.equal(view.container.querySelector('em').textContent, 'parsed');
    assert.equal(content.renderFormattedTextChildren(button, options), button);
    assert.equal(messages.filter(message => message.includes('key')).length, 0, messages.join('\n'));
  } finally {
    view?.close();
    console.error = originalError;
  }
});

test('supplied sibling keys preserve component state through reordering', () => {
  let mounts = 0;
  function Stateful({ label }) {
    const [id] = React.useState(() => ++mounts);
    return React.createElement('button', null, `${label}:${id}`);
  }
  const first = React.createElement(Stateful, { key: 'first', label: 'first' });
  const second = React.createElement(Stateful, { key: 'second', label: 'second' });
  const view = render([first, second]);
  try {
    assert.equal(view.container.textContent, 'first:1second:2');
    act(() => { ReactDOM.render(React.createElement('span', null,
      content.renderFormattedTextChildren([second, first], options)), view.container); });
    assert.equal(view.container.textContent, 'second:2first:1');
    assert.equal(mounts, 2);
  } finally { view.close(); }
});
