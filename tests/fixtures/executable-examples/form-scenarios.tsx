import * as React from 'react';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { Demo as TypedDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/4f685167cdc22f155e5c3d314de64d4911b91ebc23a5d2e0566cc9798ae92f89';
import { Demo as SchemaDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/79af64c16e48b9fef85a2e4976d0500d7c52db043b51ef3b8f341120f02dd18b';
import { Demo as DynamicDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/eeef13d46fdec7f3931e72f1ee38b4f7758c58431510d28336235f10aa0688db';
/** Presentation only: the emitted source owns methods, values and lifecycle. */
export const FormTypedScenario: React.FC = () => <TypedDemo render={regions => <DemoCard title="Edit typed reports across native fields and composition" iconName="Info" {...regions} />} />;
/** Presentation only: the emitted source owns methods, values and lifecycle. */
export const FormSchemaScenario: React.FC = () => <SchemaDemo render={regions => <DemoCard title="Validate and transform quantity with recoverable failures" iconName="Info" {...regions} />} />;
/** Presentation only: the emitted source owns methods, values and lifecycle. */
export const FormDynamicScenario: React.FC = () => <DynamicDemo render={regions => <DemoCard title="Edit a dynamic report with explicit save and reset" iconName="Info" {...regions} />} />;
/** Separate fixture routes preserve canonical scenario mounting without hidden controllers. */
export function FormProductScenario(): React.ReactElement {
  const scenario = new URLSearchParams(window.location.search).get('form-scenario');
  return scenario === 'schema' ? <FormSchemaScenario /> : scenario === 'dynamic' ? <FormDynamicScenario /> : <FormTypedScenario />;
}
