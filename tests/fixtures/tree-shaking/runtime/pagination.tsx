import * as React from 'react';
import { act } from 'react-dom/test-utils';
import { Pagination } from '__TOOLKIT__';
import type { PaginationState, PaginationChangeData } from '__TOOLKIT__';
import { FluentProvider, webDarkTheme } from '@fluentui/react-components';

export async function run({ ReactDOM, assert }: {
  ReactDOM: typeof import('react-dom'); assert: typeof import('node:assert/strict');
}) {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const ref = React.createRef<HTMLElement>();
  const requests: PaginationChangeData[] = [];
  const render = (pagination: PaginationState, loading = false): void => {
    act(() => { ReactDOM.render(<FluentProvider theme={webDarkTheme} dir="rtl">
      <Pagination ref={ref} id="installed-pagination" pagination={pagination} loading={loading}
        onPaginationChange={(_event, data) => { requests.push(data); }} />
    </FluentProvider>, container); });
  };
  const click = (label: string): void => {
    const button = container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`);
    assert.ok(button, label);
    act(() => { button.click(); });
  };
  try {
    render({ type: 'pages', pageIndex: 0, pageSize: 10, totalItems: 42 });
    assert.equal(ref.current, container.querySelector('nav'));
    click('Next page');
    assert.deepEqual(requests, [{ type: 'pages', pageIndex: 1, pageSize: 10 }]);
    assert.equal(container.querySelector<HTMLButtonElement>('button[aria-label="First page"]')?.disabled, true);
    render({ type: 'pages', pageIndex: 1, pageSize: 10, hasNextPage: true }, true);
    assert.equal(ref.current?.getAttribute('aria-busy'), 'true');
    assert.equal(container.querySelectorAll('button[disabled]').length, 2);
    click('Next page');
    assert.equal(requests.length, 1);
    render({ type: 'cursor', hasPreviousPage: true, hasNextPage: false });
    click('Previous page');
    assert.deepEqual(requests[1], { type: 'cursor', direction: 'previous' });
    render({ type: 'loadMore', hasNextPage: true });
    click('Load more');
    assert.deepEqual(requests[2], { type: 'loadMore' });
    assert.equal(container.querySelector('button')?.getAttribute('type'), 'button');
    return { assertions: ['installed native ref', 'controlled offset request', 'unknown-total loading lock', 'cursor request', 'load-more request', 'non-submit buttons'], requests };
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    container.remove();
  }
}
