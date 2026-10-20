import * as React from 'react';
import { Field, Select } from '@fluentui/react-components';
import { PnPListDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPListDemo';
import { mountDemoFixture } from '../demo-fixture-host';
mountDemoFixture(<PnPListDemo selector={busy => <Field label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'PnPPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>}><option value="list">Lista PnP</option><option value="context">Contesto</option></Select></Field>} />);
