import * as React from 'react';
import { act } from 'react-dom/test-utils';
import { EditableText } from '__TOOLKIT__';
import { EditableText as RootText } from '@apvee/spfx-react-toolkit';
import { EditableText as DomainText } from '@apvee/spfx-react-toolkit/controls';
import { EditableText as ControlText } from '@apvee/spfx-react-toolkit/controls/editable-text';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';

// The installed control and real async Tiptap run without any Shell/SPFx provider.
export async function run({ ReactDOM, assert, assetRequests }) {
  assert.equal(RootText, DomainText); assert.equal(RootText, ControlText);
  const container = document.createElement('div'); document.body.appendChild(container);
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  let firstMode: 'read' | 'edit' = 'read', secondMode: 'read' | 'edit' = 'read';
  let showFirst = true, changes = 0;
  const onChange = () => { changes++; };
  const render = async () => { await act(async () => { ReactDOM.render(
    <RendererProvider renderer={renderer} targetDocument={document}>
      {showFirst && <EditableText key="first" data-region="first" mode={firstMode} value="<h1>First</h1>" allowedBlocks={['h2', 'h4']} aria-label="First region" onChange={onChange} />}
      <EditableText key="second" data-region="second" mode={secondMode} fallbackText="Second" allowedBlocks={['h3']} aria-label="Second region" onChange={onChange} />
    </RendererProvider>, container); }); };
  const waitForEditors = async () => {
    for (let attempt = 0; attempt < 100 && container.querySelectorAll('.tiptap').length !== 2; attempt++) await act(async () => { await new Promise(resolve => setTimeout(resolve, 10)); });
    assert.equal(container.querySelectorAll('.tiptap').length, 2, 'both actual engines mount');
  };
  let first, second;
  try {
    await render();
    assert.equal(container.querySelector('[data-region="first"] h2').textContent, 'First');
    assert.equal(container.querySelector('[data-region="second"] h3').textContent, 'Second');
    assert.deepEqual(assetRequests, [], 'Read requests no editor assets'); assert.equal(changes, 0);
    firstMode = 'edit'; secondMode = 'edit'; await render(); await waitForEditors();
    first = container.querySelector('[data-region="first"] .tiptap').editor;
    second = container.querySelector('[data-region="second"] .tiptap').editor;
    assert.ok(first && second); assert.notEqual(first, second); assert.ok(assetRequests.length > 0);
    assert.equal(first.getHTML(), '<h2>First</h2>'); assert.equal(second.getHTML(), '<h3>Second</h3>');
    assert.equal(changes, 0, 'mount/block normalization never publishes');
    await act(async () => { first.commands.insertContent(' changed'); });
    assert.match(first.getText(), /changed/); assert.equal(second.getText(), 'Second');
    firstMode = 'read'; await render();
    assert.equal(first.isDestroyed, true); assert.equal(second.isDestroyed, false);
    showFirst = false; await render(); assert.equal(second.isDestroyed, false);
    await act(async () => { ReactDOM.unmountComponentAtNode(container); });
    assert.equal(second.isDestroyed, true);
    return { assertions: ['three aliases share runtime identity', 'standalone normalized Read needs no engine assets or writes', 'Edit loads two real isolated Tiptap engines', 'Read/unmount destroy only their engine'], assetRequests: [...assetRequests], destroyed: [first.isDestroyed, second.isDestroyed] };
  } finally {
    await act(async () => { ReactDOM.unmountComponentAtNode(container); });
    if (first) assert.equal(first.isDestroyed, true); if (second) assert.equal(second.isDestroyed, true);
    container.remove();
  }
}
