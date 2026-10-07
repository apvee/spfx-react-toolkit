import * as React from 'react';
import { FluentProvider } from '@fluentui/react-provider';
import { webLightTheme } from '@fluentui/react-theme';
import { Example } from './direct';
import { buildOneDriveAppDataPath } from '@apvee/spfx-react-toolkit';
export const existingToolkitCapability = buildOneDriveAppDataPath('settings.json');
export function App(): React.ReactElement {
  return <FluentProvider theme={webLightTheme}><Example /></FluentProvider>;
}
