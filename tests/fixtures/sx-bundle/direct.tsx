import * as React from 'react';
import { makeStyles } from '@griffel/react';
import { typographyStyles } from '@fluentui/react-theme';
const useStyles = makeStyles({ root: { width: '100%', ...typographyStyles.body1 } });
export function Example(): React.ReactElement {
  return <div className={useStyles().root}>Equivalent minimum UI</div>;
}
