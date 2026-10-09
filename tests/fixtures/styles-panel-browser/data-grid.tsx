import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { createDOMRenderer } from '@griffel/core';
import { RendererProvider } from '@griffel/react';
import { ResponsiveTagsFixture } from '../sx-browser/data-grid-responsive-tags';
import { LoadingBoundariesFixture } from '../sx-browser/data-grid-loading-boundaries';
import DataGridPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/DataGridPanel';

/** Run the same sample panel against the explicit compiled library boundary. */
export function renderDataGridFixture(): void {
  ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)}><section data-grid-boundary="compiled"><p>Actual DataGrid sample, React 17 and Fluent navigation. Local evidence; not authenticated SharePoint validation.</p><DataGridPanel /><ResponsiveTagsFixture /><LoadingBoundariesFixture /></section></RendererProvider>, document.getElementById('root'));
}
if (new URLSearchParams(window.location.search).has('data-grid-only')) renderDataGridFixture();
window.addEventListener('beforeunload', () => { const root = document.getElementById('root'); if (root) ReactDOM.unmountComponentAtNode(root); });
