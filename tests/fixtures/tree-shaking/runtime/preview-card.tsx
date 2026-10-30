import * as React from 'react';
import { act, Simulate } from 'react-dom/test-utils';
import { PreviewCard } from '__TOOLKIT__';
import { FluentProvider, webDarkTheme } from '@fluentui/react-components';

export async function run({ ReactDOM, assert }: {
  ReactDOM: typeof import('react-dom'); assert: typeof import('node:assert/strict');
}) {
  const container = document.createElement('div'); document.body.appendChild(container);
  const rootRef = React.createRef<HTMLDivElement>(), logoRef = React.createRef<HTMLImageElement>();
  let previewCalls = 0, fallbackCalls = 0;
  const render = (src: string, orientation: 'horizontal' | 'vertical', mode: 'image' | 'callback' | 'false', alt = 'Report cover'): void => {
    act(() => { ReactDOM.render(<FluentProvider theme={webDarkTheme} dir="rtl">
      <PreviewCard ref={rootRef} title="**Literal report**" headingLevel={2} orientation={orientation}
        image={{ src, alt }} preview={mode === 'image' ? undefined : mode === 'false' ? false : {
          children: () => { previewCalls++; return undefined; },
        }} previewFallback={{ children: (Tag, props) => { fallbackCalls++; return <Tag {...props}>Unavailable</Tag>; } }}
        logo={{ as: 'img', ref: logoRef, src: '/logo.svg', alt: 'Brand' }}>
        <input aria-label="Preview note" defaultValue="Initial" />
      </PreviewCard>
    </FluentProvider>, container); });
  };
  try {
    render('/first.jpg', 'vertical', 'image');
    assert.equal(rootRef.current?.tagName, 'DIV'); assert.equal(logoRef.current?.tagName, 'IMG');
    assert.equal(container.querySelector('h2')?.textContent, '**Literal report**');
    const input = container.querySelector<HTMLInputElement>('input'); assert.ok(input); input.value = 'Changed';
    const image = container.querySelector<HTMLImageElement>('img[alt="Report cover"]'); assert.ok(image);
    act(() => { Simulate.error(image); });
    assert.equal(container.querySelector('[data-card-image-fallback]')?.getAttribute('role'), 'img');
    assert.equal(container.querySelector('[data-card-image-fallback]')?.getAttribute('aria-label'), 'Report cover');
    assert.equal(fallbackCalls, 1);
    render('/first.jpg', 'horizontal', 'image', 'Changed alt');
    assert.equal(container.querySelector('img[alt="Changed alt"]'), null);
    assert.equal(container.querySelector('[data-card-image-fallback]')?.getAttribute('aria-label'), 'Changed alt');
    render('/second.jpg', 'horizontal', 'image');
    assert.ok(container.querySelector('img[alt="Report cover"]'));
    render('/second.jpg', 'horizontal', 'callback');
    assert.ok(container.querySelector('[data-card-preview]'));
    assert.equal(container.querySelector('img[alt="Report cover"]'), null);
    assert.equal(previewCalls, 1); assert.equal(fallbackCalls, 2);
    assert.equal(logoRef.current?.tagName, 'IMG');
    render('/second.jpg', 'vertical', 'false');
    assert.equal(container.querySelector('[data-card-preview]'), null);
    assert.equal(logoRef.current, null);
    assert.equal(container.querySelector('input'), input); assert.equal(input.value, 'Changed');
    assert.equal(rootRef.current?.hasAttribute('style'), false);
    return { assertions: ['installed native root/logo refs', 'literal heading', 'informative image failure semantics',
      'same-src failure retention and changed-src reset', 'opaque empty callback requests region and suppresses image',
      'false suppresses preview and logo', 'retained child state across preview/layout updates', 'host class rendering'] };
  } finally { act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); }
}
