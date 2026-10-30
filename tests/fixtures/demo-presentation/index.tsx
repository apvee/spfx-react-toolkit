import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { createTheme } from '@fluentui/theme';
import { type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { mountDemoFixture } from '../demo-fixture-host';
import { WebPartShellDemoHost } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/WebPartShellDemoHost';
import { fixturePage, nativeFixtureService } from './trim-transports';
import App from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/SpFxReactToolkitTest';
const element = document.getElementById('root');
if (!element) throw new Error('Fixture root missing');
// Structural host double supplies only the fields consumed by the real Shell/provider.
const theme = createTheme();
const service = { tryGetTheme: () => theme, themeChangedEvent: { add: () => {}, remove: () => {} } };
const instance = {
  instanceId: 'integrated-demo', domElement: element, properties: { title: 'Integrated toolkit demo', showTitle: false }, displayMode: 1,
  render: () => {}, context: { instanceId: 'integrated-demo', domElement: element,
    pageContext: fixturePage, sdks: {},
    serviceScope: { whenFinished: (callback: () => void) => callback(), consume: (key: unknown) => nativeFixtureService(key, service) }, propertyPane: { refresh: () => {} } },
} as unknown as BaseClientSideWebPart<IWebPartShellProperties>;
mountDemoFixture(<WebPartShellDemoHost instance={instance}><App /></WebPartShellDemoHost>);
