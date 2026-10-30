import { makeStyles } from '@griffel/react';
import { cardCSSVars, tokens } from '@fluentui/react-components';

/** Fluent owns focus/motion and the 44px native trigger minimum; Card owns lateral inset. */
export const useAccordionGroupStyles = makeStyles({
  root: { boxSizing: 'border-box', '&[hidden]': { display: 'none' } },
  card: { minInlineSize: 0 },
  accordion: {
    '& > [data-accordion-group-item] + [data-accordion-group-item]': {
      borderBlockStartWidth: tokens.strokeWidthThin,
      borderBlockStartStyle: 'solid',
      borderBlockStartColor: tokens.colorNeutralStroke2,
      // Split one Card padding interval around the divider; inherit the native surface size.
      marginBlockStart: `calc(var(${cardCSSVars.cardSizeVar}) / 2)`,
      paddingBlockStart: `calc(var(${cardCSSVars.cardSizeVar}) / 2)`,
    },
  },
  item: { '&[hidden]': { display: 'none' } },
  // Match Fluent's physical atomic properties so slot class merging removes its second inset.
  button: { paddingLeft: 0, paddingRight: 0 },
  panel: { marginLeft: 0, marginRight: 0 },
  panelContent: { paddingBlock: tokens.spacingVerticalS },
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
