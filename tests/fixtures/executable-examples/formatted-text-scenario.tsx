import * as React from 'react';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { Demo as FormattedTextDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/5ef236bf57824833f0d5ebbab7cc83ba6a3797ed10bd6b9dc2bc8d272fc1c5d9';
/** Presentation only; the emitted source owns all scenario state and actions. */
export const FormattedTextScenario: React.FC = () => <FormattedTextDemo render={regions => <DemoCard title="Explore semantic text and inline Markdown" iconName="Info" {...regions} />} />;
