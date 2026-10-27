import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem } from '@fluentui/react-components';
import ProductPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/DescriptionListPanel';
import DescriptionListPanel, { DescriptionListExample } from './description-list-regressions';
import { DemoAccordionPanel } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';

/** Independent RTL coverage is a dedicated fixture, not a duplicate product example. */
const LegacyDescriptionListIsolationFixture: React.FC = () => <>
  <DescriptionListPanel />
  <Accordion collapsible><AccordionItem value="secondary"><AccordionHeader>Independent RTL instance</AccordionHeader><DemoAccordionPanel>
    <DescriptionListExample instance="secondary" initialRtl />
  </DemoAccordionPanel></AccordionItem></Accordion>
</>;

/** Product and retained acceptance scenarios are separate executable routes. */
export const DescriptionListIsolationFixture: React.FC = () => new URLSearchParams(window.location.search).has('product-demo') ? <ProductPanel /> : <LegacyDescriptionListIsolationFixture />;
