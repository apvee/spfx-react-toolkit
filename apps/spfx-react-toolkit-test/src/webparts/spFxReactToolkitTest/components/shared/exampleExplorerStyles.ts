import { makeStyles, tokens } from '@fluentui/react-components';

/** Container queries also handle a narrow SharePoint column in a wide viewport. */
export const useExampleExplorerStyles = makeStyles({
  root: { containerType: 'inline-size', minWidth: 0 },
  layout: {
    display: 'grid', gridTemplateColumns: 'minmax(180px, 240px) minmax(0, 1fr)', gap: tokens.spacingHorizontalXXL,
    '@container (max-width: 600px)': { gridTemplateColumns: 'minmax(0, 1fr)' },
  },
  sidebar: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, minWidth: 0,
    '@container (max-width: 600px)': { display: 'none' } },
  mobile: { display: 'none', minWidth: 0, '@container (max-width: 600px)': { display: 'flex' } },
  selectRoot: { minWidth: 0, maxWidth: '100%', width: '100%' },
  selectInput: { minWidth: 0, maxWidth: '100%', width: '100%' },
  group: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, minWidth: 0 },
  groupHeader: { justifyContent: 'flex-start', textAlign: 'start', minWidth: 0, maxWidth: '100%' },
  groupLabel: { display: 'flex', flexDirection: 'column', minWidth: 0, overflow: 'hidden' },
  featureList: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, minWidth: 0, paddingInlineStart: tokens.spacingHorizontalS },
  feature: { justifyContent: 'flex-start', textAlign: 'start', minWidth: 0, maxWidth: '100%', width: '100%' },
  truncated: { overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis', minWidth: 0 },
  panel: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, minWidth: 0 },
  heading: { margin: 0, overflowWrap: 'anywhere' },
  breadcrumb: { color: tokens.colorNeutralForeground2, overflowWrap: 'anywhere' },
  toolbar: { display: 'flex', flexWrap: 'wrap', alignItems: 'flex-end', gap: tokens.spacingHorizontalM, minWidth: 0,
    '& > *': { minWidth: 0, maxWidth: '100%' } },
});
