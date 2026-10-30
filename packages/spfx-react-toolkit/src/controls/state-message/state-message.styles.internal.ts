import { makeStyles } from '@griffel/react';

/** Only details not covered by the toolkit layout descriptors. */
export const useStateMessageStyles = makeStyles({
  root: { boxSizing: 'border-box', minBlockSize: '100%' },
  visual: {
    maxInlineSize: '100%',
    '& > svg': { display: 'block', inlineSize: '100%', blockSize: 'auto', maxInlineSize: '100%' },
  },
});
