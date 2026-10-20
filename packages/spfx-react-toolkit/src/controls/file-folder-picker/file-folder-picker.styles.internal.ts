import { makeStyles, shorthands } from '@griffel/react';
import { tokens } from '@fluentui/react-theme';

/** Local descriptors use the owning host renderer; no renderer or CSS is created at module evaluation. */
export const useFileFolderPickerStyles = makeStyles({
  root: { minWidth: 0, maxWidth: '100%' },
  region: { minWidth: 0, maxWidth: '100%' },
  control: { display: 'flex', width: '100%', minWidth: 0, maxWidth: '100%', paddingTop: tokens.spacingVerticalXXS, paddingBottom: tokens.spacingVerticalXXS, paddingLeft: tokens.spacingHorizontalSNudge, paddingRight: tokens.spacingHorizontalXXS },
  selection: { flexGrow: 1, minWidth: 0 },
  placeholder: {
    display: 'block', width: '100%', ...shorthands.border('none'), backgroundColor: tokens.colorTransparentBackground,
    color: tokens.colorNeutralForeground4, fontFamily: 'inherit', fontSize: 'inherit', lineHeight: 'inherit', textAlign: 'start',
    paddingTop: 0, paddingBottom: 0, paddingLeft: tokens.spacingHorizontalSNudge, paddingRight: 0, cursor: 'pointer',
    ':focus-visible': { outlineWidth: 0 }, ':disabled': { color: tokens.colorNeutralForegroundDisabled, cursor: 'not-allowed' },
  },
  action: { flexShrink: 0 },
  open: {
    '::after': { transform: 'scaleX(1)', transitionDuration: tokens.durationNormal },
    '@media screen and (prefers-reduced-motion: reduce)': { '::after': { transitionDuration: '0.01ms', transitionDelay: '0.01ms' } },
  },
  tags: { display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS, minWidth: 0, maxWidth: '100%' },
  tag: { minWidth: 0, maxWidth: '100%' },
  text: { minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  dialog: {
    width: '1080px', height: '720px', rowGap: 0,
    // Override both native physical padding and Stage's logical padding without changing other Stage consumers.
    '&.fui-DialogSurface': { ...shorthands.padding(0), paddingBlockStart: 0, paddingBlockEnd: 0, paddingInlineStart: 0, paddingInlineEnd: 0 },
  },
  hiddenHeader: { display: 'none' },
  dialogBody: { overflow: 'hidden' },
  frame: { width: '100%', height: '100%', minHeight: 0, borderTopWidth: 0, borderRightWidth: 0, borderBottomWidth: 0, borderLeftWidth: 0, display: 'block' },
});
