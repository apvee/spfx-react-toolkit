import * as React from 'react';
import { Demo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/107f2fbceec7068e18bb4629cbda835e78ca05ada23e802ab1073c7bd29cf798';
import { mountDemoFixture } from '../demo-fixture-host';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
mountDemoFixture(<Demo render={regions => <DemoCard title="Manage list items with paging and captured delete consent" contentPlacement="configuration" {...regions} />} />);
