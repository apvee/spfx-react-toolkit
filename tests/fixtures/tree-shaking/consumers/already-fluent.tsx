import * as React from 'react';
import { FluentProvider } from '@fluentui/react-provider';
import { webLightTheme } from '@fluentui/react-theme';
import { useSx } from '__HOOK__';
import { width, typography } from '__DESCRIPTORS__';
function Child(): React.ReactElement { return <div className={useSx()(width.full, typography.body1)}>Equivalent minimum UI</div>; }
export function App(): React.ReactElement { return <FluentProvider theme={webLightTheme}><Child /></FluentProvider>; }
