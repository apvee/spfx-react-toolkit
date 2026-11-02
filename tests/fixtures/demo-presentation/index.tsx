import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { createTheme } from '@fluentui/theme';
import { type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { mountDemoFixture } from '../demo-fixture-host';
import { WebPartShellDemoHost } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/WebPartShellDemoHost';
import { fixturePage, nativeFixtureService } from './trim-transports';
import App from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/SpFxReactToolkitTest';
import { ExampleExplorer } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/ExampleExplorer';
import { executableGroups } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/catalogue';
const { filterPilotGroups } = require('./jsdoc-pilots.cjs');
const element = document.getElementById('root');
if (!element) throw new Error('Fixture root missing');
// Structural host double supplies only the fields consumed by the real Shell/provider.
const theme = createTheme();
const service = { tryGetTheme: () => theme, themeChangedEvent: { add: () => {}, remove: () => {} } };
const nativeInstance = {
  instanceId: 'integrated-demo', domElement: element, properties: { title: 'Integrated toolkit demo', showTitle: false }, displayMode: 1,
  render: () => {}, context: { instanceId: 'integrated-demo', domElement: element,
    pageContext: fixturePage, sdks: {},
    serviceScope: { whenFinished: (callback: () => void) => callback(), consume: (key: unknown) => nativeFixtureService(key, service) }, propertyPane: { refresh: () => {} } },
};
const instance = nativeInstance as unknown as BaseClientSideWebPart<IWebPartShellProperties>;
declare global { interface Window { shellExampleFixture: { setMode(mode: 'Read' | 'Edit'): void }; } }
function ShellFixtureRoot(): React.ReactElement {
  const [, invalidate] = React.useState(0);
  React.useEffect(() => {
    window.shellExampleFixture = { setMode(mode) { nativeInstance.displayMode = mode === 'Edit' ? 2 : 1; invalidate(value => value + 1); } };
  }, []);
  return <WebPartShellDemoHost instance={instance}>{(process.env.DEMO_PRESENTATION_JSDOC_ONLY === '1' || new URLSearchParams(window.location.search).get('jsdocOnly') === '1') ? <ExampleExplorer groups={filterPilotGroups(executableGroups)} legacyEntries={[]} /> : <App />}</WebPartShellDemoHost>;
}
mountDemoFixture(<ShellFixtureRoot />);
