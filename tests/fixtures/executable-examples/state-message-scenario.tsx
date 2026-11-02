import * as React from 'react';
import { Demo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/94e1205d763e4c8c9a56f32767e20121cab006a5a851417917a59b7f58fee16c';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
/** Fixture presentation only; the emitted source owns all scenario state and actions. */
export const StateMessageScenario: React.FC = () => <Demo render={regions => <DemoCard title="Configure state artwork semantics and actions" iconName="Info" {...regions} />} />;
