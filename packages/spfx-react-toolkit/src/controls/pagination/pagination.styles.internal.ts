import { makeStyles } from '@griffel/react';
import { tokens } from '@fluentui/react-components';

/** Component-only selectors and geometry not covered by layout descriptors. */
export const usePaginationStyles = makeStyles({
  root: { minBlockSize: tokens.spacingVerticalXXL, '&[hidden]': { display: 'none' } },
  icon: { ':dir(rtl)': { transform: 'scaleX(-1)' } },
});
