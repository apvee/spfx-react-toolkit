import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { createTheme } from '@fluentui/theme';
import { type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { mountDemoFixture } from '../demo-fixture-host';
import { WebPartShellDemoHost } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/WebPartShellDemoHost';
import { fixturePage, nativeFixtureService, contextRequests } from './trim-transports';
import App from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/SpFxReactToolkitTest';
import { ExampleExplorer } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/ExampleExplorer';
import { executableGroups } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/catalogue';
const { filterPilotGroups } = require('./jsdoc-pilots.cjs');
const element = document.getElementById('root');
if (!element) throw new Error('Fixture root missing');
// Structural host double supplies only the fields consumed by the real Shell/provider.
let theme: ReturnType<typeof createTheme> | undefined = createTheme();
const themeHandlers = new Set<(args: { theme?: ReturnType<typeof createTheme> }) => void>();
const service = { tryGetTheme: () => theme, themeChangedEvent: { add: (_owner: unknown, handler: (args: { theme?: ReturnType<typeof createTheme> }) => void) => themeHandlers.add(handler), remove: (_owner: unknown, handler: (args: { theme?: ReturnType<typeof createTheme> }) => void) => themeHandlers.delete(handler) } };
const nativeInstance = {
  instanceId: 'integrated-demo', domElement: element, properties: { title: 'Integrated toolkit demo', showTitle: false }, displayMode: 1,
  render: () => {}, context: { instanceId: 'integrated-demo', domElement: element,
    pageContext: fixturePage, sdks: { microsoftTeams: undefined as unknown },
    serviceScope: { whenFinished: (callback: () => void) => callback(), consume: (key: unknown) => nativeFixtureService(key, service, nativeInstance.context.pageContext) }, propertyPane: { refresh: () => {} } },
};
const instance = nativeInstance as unknown as BaseClientSideWebPart<IWebPartShellProperties>;
declare global { interface Window {
  shellExampleFixture: { setMode(mode: 'Read' | 'Edit'): void };
  contextFixture: { readonly page: typeof fixturePage; readonly requests: typeof contextRequests; rerender(): void; setPage(page: typeof fixturePage, render?: boolean): void; setSdk(sdk: unknown): void; setTheme(theme: ReturnType<typeof createTheme> | undefined): void };
} }
function ShellFixtureRoot(): React.ReactElement {
  const [, invalidate] = React.useState(0);
  React.useEffect(() => {
    window.shellExampleFixture = { setMode(mode) { nativeInstance.displayMode = mode === 'Edit' ? 2 : 1; invalidate(value => value + 1); } };
    window.contextFixture = {
      get page() { return nativeInstance.context.pageContext; }, get requests() { return contextRequests; },
      rerender: () => invalidate(value => value + 1),
      setPage(page, render = true) {
        nativeInstance.context.pageContext = page;
        nativeInstance.context.serviceScope = { whenFinished: callback => callback(), consume: key => nativeFixtureService(key, service, page) };
        if (render) invalidate(value => value + 1);
      },
      setSdk(sdk) { nativeInstance.context.sdks.microsoftTeams = sdk; invalidate(value => value + 1); },
      setTheme(next) { theme = next; for (const handler of themeHandlers) handler({ theme }); },
    };
  }, []);
  return <WebPartShellDemoHost instance={instance}>{(process.env.DEMO_PRESENTATION_JSDOC_ONLY === '1' || new URLSearchParams(window.location.search).get('jsdocOnly') === '1') ? <ExampleExplorer groups={filterPilotGroups(executableGroups)} /> : <App />}</WebPartShellDemoHost>;
}
mountDemoFixture(<ShellFixtureRoot />);
