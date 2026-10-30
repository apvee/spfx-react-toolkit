import * as React from 'react';
import { mountDemoFixture } from '../demo-fixture-host';
import { ResponsiveTagsFixture } from '../sx-browser/data-grid-responsive-tags';
import { LoadingBoundariesFixture } from '../sx-browser/data-grid-loading-boundaries';
import DataGridPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/DataGridPanel';
import RegressionPanel from '../demo-control-isolation/data-grid-regressions';

/** Run the same sample panel against the explicit compiled library boundary. */
export function renderDataGridFixture(): void {
  mountDemoFixture(<><section data-grid-boundary="compiled"><p>Actual DataGrid sample, React 17 and Fluent navigation. Local evidence; not authenticated SharePoint validation.</p>{new URLSearchParams(window.location.search).has('product-demo') ? <DataGridPanel /> : <RegressionPanel />}<ResponsiveTagsFixture /><LoadingBoundariesFixture /></section></>);
}
if (new URLSearchParams(window.location.search).has('data-grid-only')) renderDataGridFixture();
