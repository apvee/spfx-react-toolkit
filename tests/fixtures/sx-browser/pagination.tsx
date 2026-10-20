import * as React from 'react';
import { mountDemoFixture } from '../demo-fixture-host';
import { PaginationIsolationFixture } from '../demo-control-isolation/pagination';

const root = document.getElementById('root');
if (!root) throw new Error('Pagination fixture root missing');
document.title = 'Pagination source browser checkpoint';
mountDemoFixture(<>
  <PaginationIsolationFixture />
  <p data-pagination-fixture="boundary">source</p>
</>);
