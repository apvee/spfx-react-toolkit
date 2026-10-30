import * as React from 'react';
import { Accordion, AccordionItem, AccordionHeader, Card, FluentProvider } from '@fluentui/react-components';
import ProductPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StagePanel';
import StagePanel, { StageExample } from './stage-regressions';
import { DemoAccordionPanel } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';

/** A second provider and selectable ancestor exercise surface ownership independently. */
const LegacyStageIsolationFixture: React.FC = () => {
  const [selected, setSelected] = React.useState(false);
  return <>
    <StagePanel />
    <Accordion collapsible><AccordionItem value="independent"><AccordionHeader>Independent RTL instance and selectable Card</AccordionHeader><DemoAccordionPanel>
      <FluentProvider dir="rtl"><Card data-stage-demo="card" selected={selected} onSelectionChange={(_, data) => setSelected(data.selected)}>
        <p>Selectable ancestor — Stage heading must label its own surface independently.</p>
        <output data-stage-demo="card-selection">Card selected: {String(selected)}</output>
        <StageExample instance="B" />
      </Card></FluentProvider>
    </DemoAccordionPanel></AccordionItem></Accordion>
  </>;
};

/** Product and retained acceptance scenarios are separate executable routes. */
export const StageIsolationFixture: React.FC = () => new URLSearchParams(window.location.search).has('product-demo') ? <ProductPanel /> : <LegacyStageIsolationFixture />;
