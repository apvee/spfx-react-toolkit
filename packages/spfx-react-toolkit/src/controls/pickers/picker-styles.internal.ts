import { makeStyles } from '@griffel/react';
import { tokens } from '@fluentui/react-theme';

/** Uses the host renderer; no CSS insertion or renderer creation at module evaluation. */
export const usePickerStyles = makeStyles({
  root: { minWidth: 0 },
  control: { width: '100%', minWidth: 0, maxWidth: '100%' },
  persona: { minWidth: 0, flex: '1 1 0%', gridTemplateColumns: 'max-content [middle] minmax(0, 1fr)' },
  tagGroup: { minWidth: 0, maxWidth: '100%' },
  tag: { minWidth: 0, maxWidth: '100%', gridTemplateColumns: 'auto minmax(0, 1fr) auto' },
  text: { maxWidth: '100%', boxSizing: 'border-box', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', minWidth: 0 },
  status: { padding: tokens.spacingHorizontalS, whiteSpace: 'normal' },
  actions: { display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS, marginTop: tokens.spacingVerticalXS },
  list: { maxHeight: 'min(320px, 80vh)', overflowY: 'auto' },
});
