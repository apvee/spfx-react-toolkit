import * as React from 'react';
import { act } from 'react-dom/test-utils';
import { DescriptionList } from '__TOOLKIT__';
import { FluentProvider, webDarkTheme } from '@fluentui/react-components';

export async function run({ ReactDOM, assert }: {
  ReactDOM: typeof import('react-dom'); assert: typeof import('node:assert/strict');
}) {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const ref = React.createRef<HTMLDListElement>();
  const itemRef = React.createRef<HTMLDivElement>();
  let clicks = 0;
  const render = (layout: 'horizontal' | 'vertical'): void => {
    act(() => { ReactDOM.render(<FluentProvider theme={webDarkTheme} dir="rtl">
      <DescriptionList ref={ref} id="installed-description-list" layout={layout}>
        <DescriptionList.Item ref={itemRef} label="**Owner**" onClick={() => { clicks++; }}>
          <input aria-label="Owner note" defaultValue="Initial" />
        </DescriptionList.Item>
        <DescriptionList.Item label="Count">{0}</DescriptionList.Item>
      </DescriptionList>
    </FluentProvider>, container); });
  };
  try {
    render('horizontal');
    assert.equal(ref.current, container.querySelector('dl'));
    assert.equal(itemRef.current?.parentElement, ref.current);
    assert.equal(container.querySelector('dt')?.textContent, '**Owner**');
    assert.equal(container.querySelector('dt strong'), null);
    assert.equal(container.querySelectorAll('dd')[1].textContent, '0');
    const input = container.querySelector<HTMLInputElement>('input');
    assert.ok(input);
    input.value = 'Changed';
    act(() => { input.click(); });
    render('vertical');
    assert.equal(container.querySelector('input'), input);
    assert.equal(input.value, 'Changed');
    assert.equal(clicks, 1);
    assert.equal(ref.current?.hasAttribute('style'), false);
    return { assertions: ['installed native root/item refs', 'literal semantics and zero', 'native event forwarding', 'retained caller state across layouts', 'class-based host rendering'] };
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    container.remove();
  }
}
