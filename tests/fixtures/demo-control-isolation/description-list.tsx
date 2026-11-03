import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem } from '@fluentui/react-components';
import { Demo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/9ecb7aa0e0a35bac5eb68d8c09b02af3744c9b9d99a8c170af24f75e136c5d2a';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
const ProductPanel: React.FC = () => <section id="description-list-demo" data-description-list-demo="panel"><Demo render={regions => <DemoCard title="Inspect responsive metadata with data and compound APIs" iconName="Info" {...regions} />} /></section>;
import DescriptionListPanel, { DescriptionListExample } from './description-list-regressions';
import { DemoAccordionPanel } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoAccordionPanel';

/** Independent RTL coverage is a dedicated fixture, not a duplicate product example. */
const LegacyDescriptionListIsolationFixture: React.FC = () => <>
  <DescriptionListPanel />
  <Accordion collapsible><AccordionItem value="secondary"><AccordionHeader>Independent RTL instance</AccordionHeader><DemoAccordionPanel>
    <DescriptionListExample instance="secondary" initialRtl />
  </DemoAccordionPanel></AccordionItem></Accordion>
</>;

/** Product and retained acceptance scenarios are separate executable routes. */
export const DescriptionListIsolationFixture: React.FC = () => new URLSearchParams(window.location.search).has('product-demo') ? <ProductPanel /> : <LegacyDescriptionListIsolationFixture />;
