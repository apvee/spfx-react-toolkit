const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, ReactDOM, act, Pagination, FluentProvider, webLightTheme, RendererProvider, renderer, dom;
test.before(async () => {
  dom = harness.createSxDom();
  React = require('react'); ReactDOM = require('react-dom'); ({ act } = require('react-dom/test-utils'));
  ({ FluentProvider, webLightTheme } = require('@fluentui/react-components'));
  ({ RendererProvider } = require('@griffel/react'));
  renderer = require('@griffel/core').createDOMRenderer(document);
  ({ Pagination } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/pagination/pagination.tsx'));
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(props, { form = false } = {}) {
  const host = document.createElement('div'); document.body.appendChild(host);
  let submissions = 0;
  const render = (next = props) => {
    props = next;
    let content = React.createElement(Pagination, props);
    if (form) content = React.createElement('form', { onSubmit(event) { event.preventDefault(); submissions++; } }, content);
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
      React.createElement(FluentProvider, { theme: webLightTheme }, content)), host); });
  };
  render();
  return { host, render, get submissions() { return submissions; }, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
const pages = (pageIndex = 0, totalItems = 100) => ({ type: 'pages', pageIndex, pageSize: 25, totalItems });
const button = (f, label) => [...f.host.querySelectorAll('button')].find(node => node.getAttribute('aria-label') === label);
const click = node => act(() => { node.click(); });
test('known total emits controlled zero-based requests and preserves page size', () => {
  const requests = [], events = [];
  const props = { pagination: pages(1), onPaginationChange(event, data) { events.push(event.type); requests.push(data); } };
  const f = mount(props, { form: true });
  try {
    for (const label of ['First page', 'Previous page', 'Next page', 'Last page']) click(button(f, label));
    assert.deepEqual(requests, [0, 0, 2, 3].map(pageIndex => ({ type: 'pages', pageIndex, pageSize: 25 })));
    assert.deepEqual(events, ['click', 'click', 'click', 'click']);
    assert.equal(f.submissions, 0);
    assert.ok([...f.host.querySelectorAll('button')].every(node => node.type === 'button'));
    // Requests do not change the application-owned state.
    click(button(f, 'Next page')); assert.equal(requests.at(-1).pageIndex, 2);
    f.render({ ...props, pagination: pages(0) });
    assert.equal(button(f, 'First page').disabled, true); assert.equal(button(f, 'Previous page').disabled, true);
    f.render({ ...props, pagination: pages(3) });
    assert.equal(button(f, 'Next page').disabled, true); assert.equal(button(f, 'Last page').disabled, true);
    f.render({ ...props, pagination: pages(0, 0) });
    assert.ok([...f.host.querySelectorAll('button')].every(node => node.disabled));
  } finally { f.close(); }
});
test('stale known page permits explicit recovery without unsolicited correction', () => {
  const requests = [];
  const f = mount({ pagination: pages(8, 40), onPaginationChange: (_, data) => requests.push(data) });
  try {
    assert.deepEqual(requests, []);
    assert.equal(button(f, 'Next page').disabled, true); assert.equal(button(f, 'Last page').disabled, true);
    click(button(f, 'Previous page')); click(button(f, 'First page'));
    assert.deepEqual(requests.map(data => data.pageIndex), [7, 0]);
  } finally { f.close(); }
});
test('unknown totals and cursor navigation expose only available requests', () => {
  const requests = [], onPaginationChange = (_, data) => requests.push(data);
  const f = mount({ pagination: { type: 'pages', pageIndex: 2, pageSize: 10, hasNextPage: true }, onPaginationChange });
  try {
    assert.equal(f.host.querySelectorAll('button').length, 2);
    click(button(f, 'Previous page')); click(button(f, 'Next page'));
    assert.deepEqual(requests, [{ type: 'pages', pageIndex: 1, pageSize: 10 }, { type: 'pages', pageIndex: 3, pageSize: 10 }]);
    f.render({ pagination: { type: 'pages', pageIndex: Number.MAX_SAFE_INTEGER, pageSize: 10, hasNextPage: true }, onPaginationChange });
    assert.equal(button(f, 'Next page').disabled, true);
    f.render({ pagination: { type: 'pages', pageIndex: 0, pageSize: 10, hasNextPage: false }, onPaginationChange });
    assert.ok([...f.host.querySelectorAll('button')].every(node => node.disabled));
    f.render({ pagination: { type: 'cursor', hasPreviousPage: false, hasNextPage: true }, onPaginationChange });
    assert.equal(button(f, 'Previous page').disabled, true); click(button(f, 'Next page'));
    f.render({ pagination: { type: 'cursor', hasPreviousPage: true, hasNextPage: false }, onPaginationChange });
    assert.equal(button(f, 'Next page').disabled, true); click(button(f, 'Previous page'));
    assert.deepEqual(requests.slice(-2), [{ type: 'cursor', direction: 'next' }, { type: 'cursor', direction: 'previous' }]);
  } finally { f.close(); }
});
test('incremental requests, loading progress and independent instances remain consumer-owned', () => {
  let calls = 0;
  const props = { pagination: { type: 'loadMore', hasNextPage: true }, onPaginationChange(_, data) { assert.deepEqual(data, { type: 'loadMore' }); calls++; } };
  const f = mount(props), other = mount({ pagination: pages(0) });
  try {
    click(button(f, 'Load more')); assert.equal(calls, 1);
    f.render({ ...props, loading: true });
    assert.equal(button(f, 'Load more').disabled, true);
    assert.ok(f.host.querySelector('[role="progressbar"][aria-label="Loading"]'));
    assert.equal(f.host.querySelector('nav').getAttribute('aria-busy'), 'true');
    click(button(f, 'Load more')); assert.equal(calls, 1);
    assert.equal(button(other, 'Next page').disabled, false);
    f.render({ ...props, pagination: { type: 'loadMore', hasNextPage: false } });
    assert.equal(button(f, 'Load more').disabled, true);
    assert.equal(f.host.querySelector('[role="progressbar"]'), null);
    for (const pagination of [pages(1), { type: 'cursor', hasPreviousPage: true, hasNextPage: true }]) {
      f.render({ pagination, loading: true });
      assert.ok([...f.host.querySelectorAll('button')].every(node => node.disabled));
      assert.ok(f.host.querySelector('[role="progressbar"][aria-label="Loading"]'));
    }
  } finally { f.close(); other.close(); }
});
test('invalid untyped states render no navigation or callback', () => {
  const invalid = [null, undefined, {}, { type: 'bad' }, { type: 'cursor', hasPreviousPage: true },
    { type: 'loadMore', hasNextPage: 'yes' }, { ...pages(), hasNextPage: false },
    { type: 'pages', pageIndex: 0, pageSize: 25 }, { ...pages(), totalItems: undefined, hasNextPage: undefined },
    { ...pages(), totalItems: null }, { ...pages(), totalItems: '100' }, { ...pages(), hasNextPage: null }];
  for (const field of ['pageIndex', 'pageSize', 'totalItems']) {
    for (const value of [NaN, Infinity, -1, 1.5, Number.MAX_SAFE_INTEGER + 1]) invalid.push({ ...pages(), [field]: value });
  }
  invalid.push({ ...pages(), pageSize: 0 });
  let calls = 0;
  const f = mount({ pagination: pages(), onPaginationChange() { calls++; } });
  try {
    for (const pagination of invalid) { f.render({ pagination, onPaginationChange() { calls++; } }); assert.equal(f.host.querySelector('nav'), null); }
    assert.equal(calls, 0);
    f.render({ pagination: pages() }); assert.ok(f.host.querySelector('nav'));
  } finally { f.close(); }
});
test('localized labels, native forwarding, refs and owned content preserve accessibility', () => {
  const ref = React.createRef();
  const f = mount({ pagination: pages(1), ref, id: 'paging', className: 'consumer', hidden: true,
    style: { marginTop: 9 }, 'data-owner': 'consumer', 'aria-label': ' Custom nav ',
    strings: { pagination: ' Pages ', nextPage: ' Avanti ', previousPage: '  ', loading: '' },
    children: 'unexpected', dangerouslySetInnerHTML: { __html: '<b id="injected">bad</b>' } });
  try {
    assert.equal(ref.current.tagName, 'NAV'); assert.equal(ref.current.id, 'paging');
    assert.equal(ref.current.hidden, true); assert.equal(ref.current.style.marginTop, '9px');
    assert.equal(ref.current.getAttribute('data-owner'), 'consumer'); assert.ok(ref.current.classList.contains('consumer'));
    assert.equal(ref.current.getAttribute('aria-label'), 'Custom nav');
    assert.equal(f.host.querySelector('[role="group"]').getAttribute('aria-label'), 'Pages');
    assert.ok(button(f, 'Avanti')); assert.ok(button(f, 'Previous page'));
    assert.equal(f.host.querySelector('#injected'), null); assert.equal(f.host.textContent.includes('unexpected'), false);
    f.render({ pagination: pages(), 'aria-labelledby': 'external-title' });
    assert.equal(f.host.querySelector('nav').getAttribute('aria-labelledby'), 'external-title');
    assert.equal(f.host.querySelector('nav').hasAttribute('aria-label'), false);
    // JSDOM 15 retains marginTop after its setter receives ''; style removal needs a browser.
    // A fresh instance below checks that the control itself emits no inline style.
  } finally { f.close(); }
});
test('cancelled captured events do not emit pagination requests', () => {
  let calls = 0;
  const f = mount({ pagination: pages(1), onClickCapture(event) { event.preventDefault(); }, onPaginationChange() { calls++; } });
  try { click(button(f, 'Next page')); assert.equal(calls, 0); } finally { f.close(); }
});
test('native Button sizes have proportional progress in page and load-more controls', () => {
  const f = mount({ pagination: pages(1) });
  try {
    const mediumClass = button(f, 'Next page').className;
    const classes = [];
    for (const [size, spinnerSize] of [['small', 'extra-tiny'], ['medium', 'tiny'], ['large', 'extra-small']]) {
      const reference = document.createElement('div');
      document.body.appendChild(reference);
      const { Spinner } = require('@fluentui/react-components');
      act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(FluentProvider, { theme: webLightTheme }, React.createElement(Spinner, { size: spinnerSize }))), reference); });
      const expectedRingClass = reference.querySelector('.fui-Spinner__spinner').className;
      act(() => { ReactDOM.unmountComponentAtNode(reference); });
      reference.remove();
      f.render({ pagination: pages(1), size, loading: true });
      classes.push(button(f, 'Next page').className);
      const ring = button(f, 'Next page').querySelector('.fui-Spinner__spinner');
      assert.equal(ring.className, expectedRingClass, `Native progress presentation for ${size}`);
      f.render({ pagination: { type: 'loadMore', hasNextPage: true }, size, loading: true });
      assert.equal(button(f, 'Load more').querySelector('.fui-Spinner__spinner').className, expectedRingClass, `load more ${size}`);
    }
    assert.equal(new Set(classes).size, 3, 'Every native size has a distinct button presentation');
    f.render({ pagination: pages(1) });
    assert.equal(button(f, 'Next page').className, mediumClass);
  } finally { f.close(); }
});

test('appearance follows native Button presentation in every mode including loading', () => {
  const { Button } = require('@fluentui/react-components');
  const f = mount({ pagination: pages(1) });
  const defaultClass = button(f, 'Next page').className;
  const variants = [pages(1), { type: 'pages', pageIndex: 1, pageSize: 25, hasNextPage: true },
    { type: 'cursor', hasPreviousPage: true, hasNextPage: true }, { type: 'loadMore', hasNextPage: true }];
  const reference = document.createElement('div'); document.body.appendChild(reference);
  try {
    for (const appearance of ['primary', 'secondary', 'outline', 'subtle', 'transparent']) {
      for (const pagination of variants) {
        for (const loading of [false, true]) {
          const incremental = pagination.type === 'loadMore';
          act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
            React.createElement(FluentProvider, { theme: webLightTheme }, React.createElement(Button, {
              type: 'button', appearance, size: 'medium', disabled: loading,
              icon: !incremental || loading ? React.createElement('span') : undefined,
            }, incremental ? 'Load more' : undefined))), reference); });
          const expectedClass = reference.querySelector('button').className;
          f.render({ pagination, appearance, loading });
          assert.ok([...f.host.querySelectorAll('button')].every(node => node.className === expectedClass),
            `${appearance}: ${pagination.type}, loading=${loading}, native Button presentation`);
          assert.equal(f.host.querySelector('nav').hasAttribute('appearance'), false, 'Appearance belongs to buttons');
          assert.equal(f.host.querySelectorAll('[role="progressbar"]').length, loading ? 1 : 0);
        }
      }
    }
    f.render({ pagination: pages(1) });
    assert.equal(button(f, 'Next page').className, defaultClass, 'Omitted appearance retains subtle');
    f.render({ pagination: pages(1), appearance: 'subtle' });
    assert.equal(button(f, 'Next page').className, defaultClass);
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(reference); }); reference.remove(); f.close();
  }
});

test('fresh navigation uses classes without repository-authored inline declarations', () => {
  const f = mount({ pagination: pages(1) });
  try { assert.equal(f.host.querySelector('nav').hasAttribute('style'), false); } finally { f.close(); }
});
test('pending page requests replace only the activated action icon with progress', () => {
  for (const label of ['First page', 'Previous page', 'Next page', 'Last page']) {
    const requests = [];
    const props = { pagination: pages(1), onPaginationChange: (_, data) => requests.push(data) };
    const f = mount(props);
    try {
      const group = f.host.querySelector('[role="group"]');
      assert.equal(group.children.length, 4);
      assert.ok([...group.children].every(node => node.tagName === 'BUTTON'), 'No separate progress slot between navigation actions');
      click(button(f, label));
      assert.equal(requests.length, 1);
      const requested = requests[0];
      f.render({ ...props, loading: true });
      const active = button(f, label), spinner = f.host.querySelector('[role="progressbar"][aria-label="Loading"]');
      assert.ok(spinner); assert.ok(active.contains(spinner), `${label} owns the pending indicator`);
      assert.equal(active.getAttribute('aria-label'), label);
      assert.ok([...f.host.querySelectorAll('button')].every(node => node.disabled));
      assert.equal(f.host.querySelector('[role="group"]').children.length, 4);
      for (const otherLabel of ['First page', 'Previous page', 'Next page', 'Last page'].filter(value => value !== label)) {
        assert.equal(button(f, otherLabel).querySelector('[role="progressbar"]'), null);
        assert.ok(button(f, otherLabel).querySelector('svg'), `${otherLabel} retains its arrow icon`);
      }
      assert.deepEqual(requests, [requested], 'Loading UI emits no additional pagination request');
    } finally { f.close(); }
  }
});
test('completed loading resets pending action and external loading falls back to Next', () => {
  const props = { pagination: pages(1) };
  const f = mount(props);
  try {
    click(button(f, 'Last page'));
    f.render({ ...props, loading: true });
    assert.ok(button(f, 'Last page').querySelector('[role="progressbar"]'));
    f.render({ ...props, loading: false });
    assert.equal(f.host.querySelector('[role="progressbar"]'), null);
    f.render({ ...props, loading: true });
    assert.ok(button(f, 'Next page').querySelector('[role="progressbar"]'));
    assert.equal(button(f, 'Last page').querySelector('[role="progressbar"]'), null);
  } finally { f.close(); }
  for (const pagination of [pages(1), { type: 'pages', pageIndex: 1, pageSize: 25, hasNextPage: true }, { type: 'cursor', hasPreviousPage: true, hasNextPage: true }, { type: 'loadMore', hasNextPage: true }]) {
    const external = mount({ pagination, loading: true });
    try {
      const expected = pagination.type === 'loadMore' ? 'Load more' : 'Next page';
      assert.ok(button(external, expected).querySelector('[role="progressbar"][aria-label="Loading"]'));
      const group = external.host.querySelector('[role="group"]');
      assert.equal(group.children.length, pagination.type === 'loadMore' ? 1 : pagination.type === 'pages' && pagination.totalItems !== undefined ? 4 : 2);
      assert.ok([...group.children].every(node => node.tagName === 'BUTTON'));
    } finally { external.close(); }
  }
});
test('mode changes discard pending actions and instance activity stays isolated', () => {
  const f = mount({ pagination: pages(1) });
  const other = mount({ pagination: pages(1), loading: true });
  try {
    click(button(f, 'First page'));
    f.render({ pagination: pages(1), loading: true });
    assert.ok(button(f, 'First page').querySelector('[role="progressbar"]'));
    assert.ok(button(other, 'Next page').querySelector('[role="progressbar"]'));
    f.render({ pagination: { type: 'cursor', hasPreviousPage: true, hasNextPage: true }, loading: true });
    assert.ok(button(f, 'Next page').querySelector('[role="progressbar"]'));
    assert.equal(button(f, 'Previous page').querySelector('[role="progressbar"]'), null);
    assert.equal(f.host.querySelector('[role="group"]').children.length, 2);
    f.render({ pagination: { type: 'loadMore', hasNextPage: true }, loading: true });
    assert.ok(button(f, 'Load more').querySelector('[role="progressbar"]'));
    f.render({ pagination: pages(1), loading: true });
    assert.ok(button(f, 'Next page').querySelector('[role="progressbar"]'), 'Returning to pages does not revive its prior action');
  } finally { f.close(); other.close(); }
});
test('cancelled user actions do not retain a pending button indicator', () => {
  let requests = 0;
  const props = { pagination: pages(1), onClickCapture(event) { event.preventDefault(); }, onPaginationChange() { requests++; } };
  const f = mount(props);
  try {
    click(button(f, 'Last page'));
    f.render({ ...props, loading: true });
    assert.equal(requests, 0);
    assert.ok(button(f, 'Next page').querySelector('[role="progressbar"]'));
    assert.equal(button(f, 'Last page').querySelector('[role="progressbar"]'), null);
  } finally { f.close(); }
});

test('cursor pending requests place progress in the requested direction', () => {
  const props = { pagination: { type: 'cursor', hasPreviousPage: true, hasNextPage: true } };
  for (const label of ['Previous page', 'Next page']) {
    const f = mount(props);
    try {
      click(button(f, label));
      f.render({ ...props, loading: true });
      assert.ok(button(f, label).querySelector('[role="progressbar"][aria-label="Loading"]'));
      const otherLabel = label === 'Previous page' ? 'Next page' : 'Previous page';
      assert.equal(button(f, otherLabel).querySelector('[role="progressbar"]'), null);
      assert.ok(button(f, otherLabel).querySelector('svg'));
      assert.ok([...f.host.querySelector('[role="group"]').children].every(node => node.tagName === 'BUTTON'));
    } finally { f.close(); }
  }
});

test('unknown total falls back to Next if the pending first or last action disappears', () => {
  for (const label of ['First page', 'Last page']) {
    const f = mount({ pagination: pages(1) });
    try {
      click(button(f, label));
      f.render({ pagination: { type: 'pages', pageIndex: 1, pageSize: 25, hasNextPage: true }, loading: true });
      assert.ok(button(f, 'Next page').querySelector('[role="progressbar"][aria-label="Loading"]'));
      assert.equal(f.host.querySelector('[role="group"]').children.length, 2);
    } finally { f.close(); }
  }
});
test('Load more has no idle icon and shows only progress during loading', () => {
  const props = { pagination: { type: 'loadMore', hasNextPage: true } };
  const f = mount(props);
  try {
    const action = button(f, 'Load more');
    assert.equal(Boolean(action.querySelector('.fui-Button__icon')), false, 'No idle icon or empty icon slot');
    assert.equal(action.textContent, 'Load more');
    click(action);
    f.render({ ...props, loading: true });
    const pending = button(f, 'Load more');
    assert.ok(pending.querySelector('[role="progressbar"][aria-label="Loading"]'));
    assert.equal(pending.querySelector('svg'), null, 'Progress replaces the arrow');
    assert.equal(pending.textContent, 'Load more');
    assert.equal(pending.getAttribute('aria-label'), 'Load more');
    f.render(props);
    assert.equal(Boolean(button(f, 'Load more').querySelector('.fui-Button__icon')), false);
    assert.equal(f.host.querySelector('[role="progressbar"]'), null);
  } finally { f.close(); }
});


test('host-native Italian labels support trimmed partial and complete instance dictionaries', async () => {
  const { Pagination: ItalianPagination } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/pagination/pagination.tsx', { nativeCulture: 'it-IT' });
  const english = Pagination;
  Pagination = ItalianPagination;
  const italian = mount({ pagination: pages(1), loading: true, strings: { nextPage: ' Avanti ', previousPage: '  ', loading: undefined } });
  Pagination = english;
  const other = mount({ pagination: pages(1), strings: { nextPage: 'Other next' } });
  try {
    assert.equal(italian.host.querySelector('nav').getAttribute('aria-label'), 'Paginazione');
    assert.ok(button(italian, 'Avanti')); assert.ok(button(italian, 'Pagina precedente'));
    assert.ok(italian.host.querySelector('[role=progressbar][aria-label="Caricamento"]'));
    assert.ok(button(other, 'Other next')); assert.ok(button(other, 'Previous page'));
    const full = { pagination: 'Custom nav', firstPage: 'Start', previousPage: 'Back', nextPage: 'Ahead', lastPage: 'End', loadMore: 'Append', loading: 'Wait' };
    Pagination = ItalianPagination;
    italian.render({ pagination: { type: 'loadMore', hasNextPage: true }, strings: full, loading: true });
    assert.ok(button(italian, 'Append')); assert.equal(italian.host.querySelector('nav').getAttribute('aria-label'), 'Custom nav');
    assert.ok(italian.host.querySelector('[role=progressbar][aria-label="Wait"]'));
  } finally { Pagination = english; italian.close(); other.close(); }
});
