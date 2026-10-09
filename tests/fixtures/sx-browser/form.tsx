import * as React from 'react';
import FormPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormPanel';

/** Exercise the actual sample against source-defined Form modules via the fixture's explicit boundary. */
export function FormFixture(): React.ReactElement {
  return <section data-form-boundary="source"><p>Source Form controls, React 17 and shared FluentProvider. This is local browser evidence, not authenticated SharePoint validation.</p><FormPanel /></section>;
}
