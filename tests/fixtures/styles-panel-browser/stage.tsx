import * as React from 'react';
import { mountDemoFixture } from '../demo-fixture-host';
import { StageIsolationFixture } from '../demo-control-isolation/stage';
import { StageBrowserProbes } from '../stage-browser-probes';

const root = document.getElementById('root');
if (!root) throw new Error('Stage fixture root missing');
document.title = 'Stage compiled browser checkpoint';
mountDemoFixture(<>
  <StageIsolationFixture />
  <StageBrowserProbes />
  <button data-stage-fixture="outside">Outside Stage</button>
  <p data-stage-fixture="boundary">compiled</p>
</>);
