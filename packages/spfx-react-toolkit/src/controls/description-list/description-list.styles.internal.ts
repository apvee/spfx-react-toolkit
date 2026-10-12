import { makeStyles } from '@griffel/react';
import { tokens } from '@fluentui/react-theme';

/** Component-owned unequal grid and query rules absent from the shared descriptor catalog. */
export const useDescriptionListStyles = makeStyles({
  root: { display: 'grid', '&[hidden]': { display: 'none' } },
  vertical: { gridTemplateColumns: 'minmax(0, 1fr)' },
  horizontal: {
    gridTemplateColumns: 'fit-content(40%) minmax(0, 1fr)',
    columnGap: tokens.spacingHorizontalM,
    rowGap: tokens.spacingVerticalS,
  },
  auto: {
    gridTemplateColumns: 'minmax(0, 1fr)',
    '@container apvee-sx (min-inline-size: 480px)': {
      gridTemplateColumns: 'fit-content(40%) minmax(0, 1fr)',
      columnGap: tokens.spacingHorizontalM,
      rowGap: tokens.spacingVerticalS,
    },
  },
  item: { '&[hidden]': { display: 'none' } },
  itemVertical: { display: 'grid', gridTemplateColumns: 'minmax(0, 1fr)', rowGap: tokens.spacingVerticalXXS },
  itemHorizontal: { display: 'contents' },
  itemAuto: {
    display: 'grid', gridTemplateColumns: 'minmax(0, 1fr)', rowGap: tokens.spacingVerticalXXS,
    '@container apvee-sx (min-inline-size: 480px)': { display: 'contents' },
  },
  cell: { overflowWrap: 'anywhere' },
});
