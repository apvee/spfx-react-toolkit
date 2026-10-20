import * as React from 'react';
import { mountDemoFixture } from '../demo-fixture-host';
import { CardsIsolationFixture } from '../demo-control-isolation/cards';
const root = document.getElementById('root');
if (!root) throw new Error('Cards fixture root missing');
document.title = 'Cards compiled browser checkpoint';
mountDemoFixture(<>
  <CardsIsolationFixture /><p data-cards-fixture="boundary">compiled</p>
</>);
