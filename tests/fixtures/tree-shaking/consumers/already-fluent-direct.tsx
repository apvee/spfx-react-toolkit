import * as React from 'react';
import { FluentProvider } from '@fluentui/react-provider';
import { webLightTheme } from '@fluentui/react-theme';
import { makeStyles } from '@griffel/react';
import { typographyStyles } from '@fluentui/react-theme';
const useStyles = makeStyles({ root: { width: '100%', ...typographyStyles.body1 } });
function Child(): React.ReactElement { return <div className={useStyles().root}>Equivalent minimum UI</div>; }
export function App(): React.ReactElement { return <FluentProvider theme={webLightTheme}><Child /></FluentProvider>; }
