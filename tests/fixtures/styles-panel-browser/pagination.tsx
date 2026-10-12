import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';
import PaginationPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/PaginationPanel';

const root = document.getElementById('root');
if (!root) throw new Error('Pagination fixture root missing');
document.title = 'Pagination compiled browser checkpoint';
ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)} targetDocument={document}>
  <PaginationPanel />
  <p data-pagination-fixture="boundary">compiled</p>
</RendererProvider>, root);
