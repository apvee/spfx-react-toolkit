import * as React from 'react';
import { Accordion, AccordionItem, AccordionHeader, FluentProvider } from '@fluentui/react-components';
import CardsPanel, { CardExamples } from './cards-regressions';
import { DemoAccordionPanel } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';
/** Native comparisons and pathological slot cases remain executable outside the product. */
export const CardsLegacyIsolationFixture: React.FC = () => <>
  <CardsPanel />
  <Accordion collapsible><AccordionItem value="independent"><AccordionHeader>Independent RTL instance</AccordionHeader><DemoAccordionPanel>
    <FluentProvider dir="rtl"><CardExamples name="independent" /></FluentProvider>
  </DemoAccordionPanel></AccordionItem></Accordion>
</>;
