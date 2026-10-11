import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { createDOMRenderer } from '@griffel/core';
import { RendererProvider } from '@griffel/react';
import FormPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormPanel';

/** Run the actual Form sample against compiled library modules without the SPFx graph. */
ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)}>
  <section data-form-boundary="compiled"><FormPanel /></section>
</RendererProvider>, document.getElementById('root'));
window.addEventListener('beforeunload', () => {
  const root = document.getElementById('root');
  if (root) ReactDOM.unmountComponentAtNode(root);
});
