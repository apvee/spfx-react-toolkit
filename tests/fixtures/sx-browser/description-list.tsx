import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';
import DescriptionListPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/DescriptionListPanel';
import { DescriptionListProbes } from '../description-list-browser/probes';
const root = document.getElementById('root');
if (!root) throw new Error('DescriptionList fixture root missing');
document.title = 'DescriptionList source browser checkpoint';
ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)} targetDocument={document}>
  <DescriptionListPanel />
  <DescriptionListProbes />
  <p data-description-list-fixture="boundary">source</p>
</RendererProvider>, root);
