import * as React from 'react';
import { mountDemoFixture } from '../demo-fixture-host';
import { FormProductScenario } from '../executable-examples/form-scenarios';
import RegressionPanel from '../demo-control-isolation/form-regressions';

/** Run the actual Form sample against compiled library modules without the SPFx graph. */
mountDemoFixture(<>
  <section data-form-boundary="compiled">{new URLSearchParams(window.location.search).has('product-demo') ? <FormProductScenario /> : <RegressionPanel />}</section>
</>);
