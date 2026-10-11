import * as React from 'react';
import { act, Simulate } from 'react-dom/test-utils';
import { Stage } from '__TOOLKIT__';
import { FluentProvider, webLightTheme, webDarkTheme } from '@fluentui/react-components';

export async function run({ ReactDOM, assert }) {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const ref = React.createRef<HTMLDivElement>();
  const requests: Array<{ open: boolean; reason: string }> = [];
  const content = <>
    <Stage.Header title="**Installed profile**" description="*Literal details*" closeLabel="Close installed profile" />
    <Stage.Body><input aria-label="Installed name" defaultValue="Initial" /></Stage.Body>
    <Stage.Footer><button type="button">Installed action</button></Stage.Footer>
  </>;
  const render = async (mode: 'dialog' | 'drawer' | 'popover', size: 'small' | 'medium' = 'small', dark = false) => {
    const common = {
      ref, open: true, size, id: 'installed-stage', unmountOnClose: false,
      onOpenChange: (_event, data) => { requests.push({ open: data.open, reason: data.reason }); },
      trigger: <button type="button">Open installed profile</button>,
    };
    await act(async () => ReactDOM.render(<FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir="rtl">
      {mode === 'popover'
        ? <Stage {...common} presentation="popover">{content}</Stage>
        : <Stage {...common} presentation={mode}>{content}</Stage>}
    </FluentProvider>, container));
  };
  try {
    await render('dialog');
    assert.equal(ref.current, document.querySelector('#installed-stage'));
    const surface = ref.current;
    assert.equal(surface.getAttribute('role'), 'dialog');
    assert.equal(surface.getAttribute('aria-modal'), 'true');
    const title = document.getElementById(surface.getAttribute('aria-labelledby'));
    assert.equal(title.textContent, '**Installed profile**');
    assert.equal(title.querySelector('strong'), null);
    assert.equal(document.getElementById(surface.getAttribute('aria-describedby')).textContent, '*Literal details*');
    const input = surface.querySelector('input');
    input.value = 'Changed';
    await render('dialog', 'medium', true);
    assert.equal(ref.current, surface);
    assert.equal(ref.current.querySelector('input'), input);
    assert.equal(input.value, 'Changed');
    const close = surface.querySelector('button[aria-label="Close installed profile"]');
    await act(async () => Simulate.click(close));
    assert.deepEqual(requests, [{ open: false, reason: 'closeButton' }]);
    assert.equal(ref.current, surface, 'controlled open stays authoritative');
    assert.ok(Stage.Header && Stage.Body && Stage.Footer);
    return {
      assertions: ['canonical compound exports', 'installed modal surface and accessible IDs', 'literal header text',
        'native root ref', 'size/theme update preserves body input', 'controlled close request'],
      requests, preservedInput: input.value,
    };
  } finally {
    await act(async () => ReactDOM.unmountComponentAtNode(container));
    container.remove();
  }
}
