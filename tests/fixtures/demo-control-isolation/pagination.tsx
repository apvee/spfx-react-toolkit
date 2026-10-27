import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem } from '@fluentui/react-components';
import ProductPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/PaginationPanel';
import PaginationPanel, { PaginationExample } from './pagination-regressions';
import { DemoAccordionPanel } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';

/** Independent RTL coverage is a dedicated fixture, not a duplicate product example. */
const LegacyPaginationIsolationFixture: React.FC = () => <>
  <PaginationPanel />
  <Accordion collapsible><AccordionItem value="secondary"><AccordionHeader>Independent RTL instance</AccordionHeader><DemoAccordionPanel>
    <PaginationExample instance="secondary" initialRtl />
  </DemoAccordionPanel></AccordionItem></Accordion>
</>;

/** Product and retained acceptance scenarios are separate executable routes. */
export const PaginationIsolationFixture: React.FC = () => new URLSearchParams(window.location.search).has('product-demo') ? <ProductPanel /> : <LegacyPaginationIsolationFixture />;
