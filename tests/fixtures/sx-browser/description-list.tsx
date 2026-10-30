import * as React from 'react';
import { mountDemoFixture } from '../demo-fixture-host';
import { DescriptionListIsolationFixture } from '../demo-control-isolation/description-list';
import { DescriptionListProbes } from '../description-list-browser/probes';
const root = document.getElementById('root');
if (!root) throw new Error('DescriptionList fixture root missing');
document.title = 'DescriptionList source browser checkpoint';
mountDemoFixture(<>
  <DescriptionListIsolationFixture />
  <DescriptionListProbes />
  <p data-description-list-fixture="boundary">source</p>
</>);
