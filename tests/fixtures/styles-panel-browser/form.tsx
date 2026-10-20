import * as React from 'react';
import { mountDemoFixture } from '../demo-fixture-host';
import FormPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormPanel';
import RegressionPanel from '../demo-control-isolation/form-regressions';

/** Run the actual Form sample against compiled library modules without the SPFx graph. */
mountDemoFixture(<>
  <section data-form-boundary="compiled">{new URLSearchParams(window.location.search).has('product-demo') ? <FormPanel /> : <RegressionPanel />}</section>
</>);
