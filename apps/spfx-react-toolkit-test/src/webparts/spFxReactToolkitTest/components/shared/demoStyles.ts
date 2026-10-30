import { makeStyles, tokens } from '@fluentui/react-components';

/** Shared demo regions inherit the host renderer, theme and direction. */
export const useDemoStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXL, minWidth: 0 },
  card: { padding: tokens.spacingVerticalL, gap: tokens.spacingVerticalL, minWidth: 0, backgroundColor: 'transparent', boxShadow: 'none' },
  section: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXL, minWidth: 0 },
  sectionHeader: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, minWidth: 0 },
  heading: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, margin: 0 },
  configTitle: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, fontSize: tokens.fontSizeBase400 },
  description: { margin: 0, color: tokens.colorNeutralForeground2 },
  controls: {
    display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    gap: tokens.spacingVerticalM, minWidth: 0,
    '& > *': { minWidth: 0, maxWidth: '100%' },
    '@media (max-width: 600px)': { gridTemplateColumns: 'minmax(0, 1fr)' },
  },
  controlRow: { gridColumn: '1 / -1' },
  preview: { minWidth: 0, maxWidth: '100%', '&:empty': { display: 'none' } },
  results: { minWidth: 0, maxWidth: '100%' },
  status: { display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: tokens.spacingHorizontalS, minWidth: 0, fontSize: tokens.fontSizeBase200, '& > p': { margin: 0 }, overflowWrap: 'anywhere' },
  provider: { minWidth: 0, backgroundColor: 'transparent' },
  actions: { display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: tokens.spacingHorizontalS, minWidth: 0 },
  observations: { minWidth: 0, overflowWrap: 'anywhere' },
  scroll: { overflowX: 'auto', minWidth: 0, maxWidth: '100%', borderRadius: tokens.borderRadiusMedium },
  tabList: { display: 'flex', flexWrap: 'wrap', minWidth: 0, '& > *': { maxWidth: '100%' } },
  panel: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, minWidth: 0 },
  infoGrid: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, minWidth: 0 },
  infoRow: { display: 'flex', flexWrap: 'wrap', alignItems: 'flex-start', gap: tokens.spacingHorizontalS, minWidth: 0, overflowWrap: 'anywhere' },
  infoLabel: { display: 'inline-flex', alignItems: 'center', gap: tokens.spacingHorizontalS, minWidth: 0, maxWidth: '100%', '& > .fui-Text': { minWidth: 0 } },
  json: { margin: 0, padding: tokens.spacingVerticalM, fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, backgroundColor: tokens.colorNeutralBackground2 },
});
