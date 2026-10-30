import * as React from 'react';
import { act } from 'react-dom/test-utils';
import { ContentCard } from '__TOOLKIT__';
import { Button, Card, FluentProvider, webDarkTheme } from '@fluentui/react-components';

export async function run({ ReactDOM, assert }: {
  ReactDOM: typeof import('react-dom'); assert: typeof import('node:assert/strict');
}) {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const rootRef = React.createRef<HTMLDivElement>(), titleRef = React.createRef<HTMLSpanElement>();
  let actions = 0, selections = 0, nativeSelections = 0, slotCalls = 0;
  const render = (disabled: boolean): void => {
    act(() => { ReactDOM.render(<FluentProvider theme={webDarkTheme} dir="rtl">
      <ContentCard data-probe="content" ref={rootRef} disabled={disabled} defaultSelected={false}
        title={{ ref: titleRef, className: 'consumer-title', children: (Tag, props) => {
          slotCalls++; return <Tag {...props}>**Literal summary**</Tag>;
        } }} headingLevel={2} onSelectionChange={() => { selections++; }}
        actions={<Button data-probe="action" onClick={() => { actions++; }}>Open</Button>}>
        <input aria-label="Content note" defaultValue="Initial" />
      </ContentCard>
      <Card data-probe="native" disabled={disabled} defaultSelected={false} onSelectionChange={() => { nativeSelections++; }}>Native peer</Card>
    </FluentProvider>, container); });
  };
  try {
    render(false);
    assert.equal(rootRef.current, container.querySelector('[data-probe="content"]'));
    assert.equal(titleRef.current?.parentElement?.tagName, 'H2');
    assert.equal(titleRef.current?.textContent, '**Literal summary**');
    assert.ok(titleRef.current?.classList.contains('consumer-title'));
    assert.equal(container.querySelector('strong'), null);
    assert.equal(slotCalls, 1);
    const note = container.querySelector<HTMLInputElement>('input[aria-label="Content note"]');
    assert.ok(note); note.value = 'Changed';
    const content = rootRef.current, native = container.querySelector<HTMLElement>('[data-probe="native"]');
    assert.ok(content); assert.ok(native);
    act(() => { content.click(); native.click(); });
    assert.equal(selections, nativeSelections);
    assert.ok(selections > 0);
    render(true);
    const action = container.querySelector<HTMLButtonElement>('[data-probe="action"]');
    assert.ok(action); assert.equal(action.disabled, false);
    act(() => { action.click(); });
    assert.equal(actions, 1);
    assert.equal(container.querySelector('input[aria-label="Content note"]'), note);
    assert.equal(note.value, 'Changed');
    assert.equal(rootRef.current?.hasAttribute('style'), false);
    return { assertions: ['installed native root and inner title refs', 'literal callback title and caller class',
      'native selection parity', 'disabled leaves caller controls enabled', 'retained child state', 'host class rendering'] };
  } finally { act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); }
}
