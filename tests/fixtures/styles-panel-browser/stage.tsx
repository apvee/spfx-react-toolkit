import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';
import StagePanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StagePanel';
import { StageBrowserProbes } from '../stage-browser-probes';

const root = document.getElementById('root');
if (!root) throw new Error('Stage fixture root missing');
document.title = 'Stage compiled browser checkpoint';
ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)} targetDocument={document}>
  <StagePanel />
  <StageBrowserProbes />
  <button data-stage-fixture="outside">Outside Stage</button>
  <p data-stage-fixture="boundary">compiled</p>
</RendererProvider>, root);
