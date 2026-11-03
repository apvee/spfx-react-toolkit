import * as React from 'react';
import { Demo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/6010887cfaf12213f5fc382a3f51f03ca038e58485625b8c904f3b8fb0f01e30';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
/** Presentation only; the emitted source owns all recipe state and actions. */
export const StylesScenario: React.FC = () => <Demo render={regions => <DemoCard title="Explore eight style recipes with the host renderer" iconName="Info" contentPlacement="configuration" {...regions} />} />;
