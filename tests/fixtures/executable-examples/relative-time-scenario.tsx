import * as React from 'react';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { Demo as RelativeTimeDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/a1fadbadc90bc687cdc11d676df4b5fae241637fb53645f576789fed564f278a';
/** Presentation only; the emitted source owns all scenario state and actions. */
export const RelativeTimeScenario: React.FC = () => <RelativeTimeDemo render={regions => <DemoCard title="Explore relative clocks and fallback slots" iconName="Info" {...regions} />} />;
