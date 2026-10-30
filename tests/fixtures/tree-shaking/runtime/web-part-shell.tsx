import * as React from 'react';
import { act } from 'react-dom/test-utils';
import { WebPartShell as Shell, getWebPartShellPropertiesMetadata } from '__TOOLKIT__';
import { WebPartShell as RootShell, getWebPartShellPropertiesMetadata as rootMetadata } from '@apvee/spfx-react-toolkit';
import { WebPartShell as DomainShell, getWebPartShellPropertiesMetadata as domainMetadata } from '@apvee/spfx-react-toolkit/controls';
import { WebPartShell as ControlShell, getWebPartShellPropertiesMetadata as controlMetadata } from '@apvee/spfx-react-toolkit/controls/web-part-shell';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';

// Actual tarball exports, provider, title, engine and renderer; only SPFx services
// and the Node/JSDOM asset delivery surface are supplied by the runtime host.
export async function run({ ReactDOM, assert, assetRequests }) {
  assert.equal(RootShell, DomainShell); assert.equal(RootShell, ControlShell);
  assert.equal(rootMetadata, domainMetadata); assert.equal(rootMetadata, controlMetadata);
  const first = getWebPartShellPropertiesMetadata(), second = getWebPartShellPropertiesMetadata();
  assert.deepEqual(first, { title: { isSearchablePlainText: true } });
  assert.notEqual(first, second); assert.notEqual(first.title, second.title);
  const container = document.createElement('div'); document.body.appendChild(container);
  const listeners = new Set(); const theme = { tryGetTheme: () => undefined, themeChangedEvent: {
    add: (_, handler) => listeners.add(handler), remove: (_, handler) => listeners.delete(handler),
  } };
  const context = { instanceId: 'installed-shell', domElement: container, pageContext: { cultureInfo: { isRightToLeft: false } },
    propertyPane: { refresh() {} }, serviceScope: { whenFinished: fn => fn(), consume: key => { assert.equal(key, 'production-theme'); return theme; } } };
  let instance = { instanceId: context.instanceId, context, domElement: container, properties: { title: 'Installed title', listId: 'keep' }, displayMode: 1, render() {} };
  let mounts = 0, cleanups = 0;
  function Child() { React.useEffect(() => { mounts++; return () => { cleanups++; }; }, []); return <span data-stable-child="">Content</span>; }
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const render = async () => { await act(async () => { ReactDOM.render(<RendererProvider renderer={renderer} targetDocument={document}>
    <Shell instance={instance}><Child /></Shell></RendererProvider>, container); }); };
  let editor;
  try {
    await render();
    assert.equal(container.querySelector('h2').textContent, 'Installed title');
    assert.equal(container.querySelector('[contenteditable]'), null);
    assert.deepEqual(assetRequests, [], 'Read must request no async engine assets');
    const child = container.querySelector('[data-stable-child]');
    instance = { ...instance, displayMode: 2 }; await render();
    for (let attempt = 0; attempt < 100 && !container.querySelector('.tiptap'); attempt++) await act(async () => { await new Promise(resolve => setTimeout(resolve, 10)); });
    const editable = container.querySelector('.tiptap'); assert.ok(editable, 'Edit loads the actual editor');
    editor = editable.editor; assert.ok(editor && !editor.isDestroyed, 'Actual Tiptap Editor owns its mounted view');
    assert.ok(assetRequests.length > 0, 'Edit requests the emitted editor chunk');
    assert.equal(editable.textContent, 'Installed title');
    assert.equal(child, container.querySelector('[data-stable-child]')); assert.equal(mounts, 1);
    instance = { ...instance, displayMode: 1 }; await render();
    assert.equal(editor.isDestroyed, true); assert.equal(container.querySelector('[contenteditable]'), null);
    assert.equal(cleanups, 0); assert.equal(child, container.querySelector('[data-stable-child]'));
    assert.deepEqual(instance.properties, { title: 'Installed title', listId: 'keep' });
    return { assertions: ['root/domain/control runtime identity', 'fresh pure metadata', 'Read no engine or chunk request', 'Edit requests emitted chunks and mounts actual Tiptap view', 'Read destroys engine and retains child without writes'], assetRequests: [...assetRequests], editorDestroyed: editor.isDestroyed, childMounts: mounts };
  } finally {
    await act(async () => { ReactDOM.unmountComponentAtNode(container); });
    assert.equal(listeners.size, 0); assert.equal(cleanups, 1); if (editor) assert.equal(editor.isDestroyed, true);
    container.remove();
  }
}
