import { makeStyles } from '@griffel/react';
import { cardCSSVars, tokens } from '@fluentui/react-components';

/** Only selectors and details absent from existing layout descriptors. */
export const useSettingsGroupStyles = makeStyles({
  root: { boxSizing: 'border-box', '&[hidden]': { display: 'none' } },
  card: {
    minInlineSize: 0,
    '& > [data-settings-group-item] + [data-settings-group-item]': {
      borderBlockStartWidth: tokens.strokeWidthThin,
      borderBlockStartStyle: 'solid',
      borderBlockStartColor: tokens.colorNeutralStroke2,
      // Split one Card padding interval around the divider; inherit the native surface size.
      marginBlockStart: `calc(var(${cardCSSVars.cardSizeVar}) / 2)`,
      paddingBlockStart: `calc(var(${cardCSSVars.cardSizeVar}) / 2)`,
    },
  },
  item: { '&[hidden]': { display: 'none' } },
  // Keep the row minimum independent of the outer divider's padding.
  row: { rowGap: tokens.spacingVerticalS },
  leading: { flexBasis: '160px', minInlineSize: 'min(100%, 160px)', maxInlineSize: '100%' },
  controls: { maxInlineSize: '100%' },
  // Only the introduced object/callback container participates in owned icon sizing.
  iconContent: {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    inlineSize: '100%', blockSize: '100%',
    '& > i': { fontSize: 'inherit' },
    '& > svg': { inlineSize: '100%', blockSize: '100%' },
  },
  icon: {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    inlineSize: '20px', blockSize: '20px', fontSize: '20px', lineHeight: '1',
    '& > i': { fontSize: 'inherit' },
    '& > svg': { inlineSize: '100%', blockSize: '100%' },
  },
});
