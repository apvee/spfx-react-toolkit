import * as React from 'react';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { Demo as EditableTextDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/e4d8e4f3dcffb7af68fa9531d2ba7c9824bb0aef10da8691df748d465b944e57';
/** Presentation only; the emitted source owns all scenario state and actions. */
export const EditableTextScenario: React.FC = () => <EditableTextDemo render={regions => <DemoCard title="Edit a controlled block with policy and dictionaries" iconName="Info" {...regions} />} />;
