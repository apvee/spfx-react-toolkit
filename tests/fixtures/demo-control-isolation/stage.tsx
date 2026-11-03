import * as React from 'react';
import { Accordion, AccordionItem, AccordionHeader, Card, FluentProvider } from '@fluentui/react-components';
import { Demo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/06e0030c1874dec9ca0a1d76b9acaf10b2d207e7c562ac7baf9ec8ac37986052';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
const ProductPanel: React.FC = () => <Demo render={regions => <DemoCard title="Edit a report with guarded Stage dismissal" iconName="Info" contentPlacement="configuration" {...regions} />} />;
import StagePanel, { StageExample } from './stage-regressions';
import { DemoAccordionPanel } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoAccordionPanel';

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
