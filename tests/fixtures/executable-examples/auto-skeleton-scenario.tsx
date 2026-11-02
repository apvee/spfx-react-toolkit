import * as React from 'react';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { Demo as AutoSkeletonDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/e075dc16007a7bb79e170fac496a3d7a9939e6882b4c4f4b0de8b18f71fceb04';
/** Presentation only; the emitted source owns all scenario state and actions. */
export const AutoSkeletonScenario: React.FC = () => <AutoSkeletonDemo render={regions => <DemoCard title="Explore mounted profile loading and native focus" iconName="Info" {...regions} />} />;
