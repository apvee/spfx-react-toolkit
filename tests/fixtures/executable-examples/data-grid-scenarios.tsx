import * as React from 'react';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { Demo as TypedDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/72da85f16f9d5f684b5c802effaf3ef8c657ada6b8ea7b629af6f08734bc954f';
import { Demo as DescriptorDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/abf494bc983f86eb12fc98b6bfad805ef6dbd7eff892bb7116062585d6142a4c';
import { Demo as DynamicDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/a20740a2c7e96b1271767e6ac6c1bf0ade4fe7f337ff4c9bf8638d01b82d3bd9';
/** Presentation only: the emitted source owns all grid data, actions and lifecycle. */
export const GridTypedScenario: React.FC = () => <TypedDemo render={regions => <DemoCard title="Explore typed document columns and controlled interaction" iconName="Info" {...regions} />} />;
/** Presentation only: the emitted source owns all grid data, actions and lifecycle. */
export const GridDescriptorScenario: React.FC = () => <DescriptorDemo render={regions => <DemoCard title="Supply descriptor columns with explicit sort and row order" iconName="Info" {...regions} />} />;
/** Presentation only: the emitted source owns all grid data, actions and lifecycle. */
export const GridDynamicScenario: React.FC = () => <DynamicDemo render={regions => <DemoCard title="Classify dynamic row values and preserve controlled selection" iconName="Info" {...regions} />} />;
/** Distinct routes mount only the selected actual source-owned scenario. */
export function GridProductScenario(): React.ReactElement {
  const scenario = new URLSearchParams(window.location.search).get('grid-scenario');
  return scenario === 'descriptors' ? <GridDescriptorScenario /> : scenario === 'dynamic' ? <GridDynamicScenario /> : <GridTypedScenario />;
}
