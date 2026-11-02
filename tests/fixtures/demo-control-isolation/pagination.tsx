import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem } from '@fluentui/react-components';
import { Demo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/a07f48c5bc9c9b2205540773cc72cbeeadb752470bb6c3e67ff22ec10c3e5ee0';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
const ProductScenario: React.FC = () => <section data-pagination-demo="panel"><Demo render={regions => <DemoCard title="Browse orders across pagination modes" iconName="Info" {...regions} />} /></section>;
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
export const PaginationIsolationFixture: React.FC = () => new URLSearchParams(window.location.search).has('product-demo') ? <ProductScenario /> : <LegacyPaginationIsolationFixture />;
