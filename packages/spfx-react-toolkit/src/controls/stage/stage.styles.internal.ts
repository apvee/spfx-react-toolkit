import { makeStyles } from '@griffel/react';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as alignItems from '../../helpers/styles/align-items';
import * as gap from '../../helpers/styles/gap';
import * as padding from '../../helpers/styles/padding';
import * as minHeight from '../../helpers/styles/min-height';
import * as minWidth from '../../helpers/styles/min-width';
import * as overflow from '../../helpers/styles/overflow';
import * as background from '../../helpers/styles/background';
import * as boxShadow from '../../helpers/styles/box-shadow';
import * as borderRadius from '../../helpers/styles/border-radius';
import * as container from '../../helpers/styles/container';
import * as width from '../../helpers/styles/width';

const dialogGap = 'clamp(8px, 1vw, 16px)';
const dialogWidth = `calc(100vw - 2 * ${dialogGap})`;
const dialogHeight = `calc(100vh - 2 * ${dialogGap})`;
const dynamicDialogHeight = `calc(100dvh - 2 * ${dialogGap})`;
const useSpecificStyles = makeStyles({
  backdrop: { '&[hidden]': { display: 'none', pointerEvents: 'none' } },
  surface: {
    boxSizing: 'border-box',
    // Let the surface scroll as a whole when fixed header/footer content would crowd Body.
    '@media screen and (max-height: 359px)': { overflowY: 'auto' },
    '&[hidden]': { display: 'none' },
  },
  header: { '&[hidden]': { display: 'none' } },
  body: {
    '@media screen and (max-height: 359px)': { flexShrink: 0, overflowY: 'visible' },
    '&[hidden]': { display: 'none' },
  },
  footer: { '&[hidden]': { display: 'none' } },
  neutralCard: { '&::after': { display: 'none' } },
  cardHeader: { display: 'grid', gridTemplateColumns: 'auto minmax(0, 1fr) auto', rowGap: 0 },
  text: { gridColumnStart: 2, gridRowStart: 1 },
  description: { gridColumnStart: 2, gridRowStart: 2 },
  icon: {
    gridColumnStart: 1, gridRowStart: 'span 2',
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    inlineSize: '24px', blockSize: '24px', fontSize: '24px', lineHeight: '1',
    '& > i': { fontSize: 'inherit' },
    '& > svg': { inlineSize: '100%', blockSize: '100%' },
  },
  actions: {
    gridColumnStart: 3, gridRowStart: 'span 2',
    // Query the stretched neutral Header, never the viewport or presentation mode.
    '@container apvee-sx (max-width: 400px)': {
      gridColumnStart: 1, gridColumnEnd: -1, gridRowStart: 3, marginLeft: 0, marginRight: 0,
    },
    maxInlineSize: '100%', overflowWrap: 'anywhere',
  },
  dialog: {
    maxWidth: dialogWidth, maxHeight: dialogHeight,
    // Native reset styles set max-width:100vw in a later media bucket. The
    // public surface class keeps the Stage gap authoritative at narrow widths.
    '&.fui-DialogSurface': { maxWidth: dialogWidth },
    '@supports (height: 100dvh)': { maxHeight: dynamicDialogHeight },
  },
  dialogSmall: { width: '640px' },
  dialogMedium: { width: '800px' },
  dialogLarge: { width: '1120px' },
  dialogFull: {
    width: dialogWidth, height: dialogHeight,
    '@supports (height: 100dvh)': { height: dynamicDialogHeight },
  },
  drawer: { maxWidth: '100vw', height: '100vh', maxHeight: '100vh',
    '@supports (height: 100dvh)': { height: '100dvh', maxHeight: '100dvh' } },
  drawerSmall: { width: '320px' },
  drawerMedium: { width: '592px' },
  drawerLarge: { width: '940px' },
  drawerFull: { width: '100vw' },
  popover: { maxWidth: '100vw', maxHeight: '100vh',
    '@supports (height: 100dvh)': { maxHeight: '100dvh' } },
  popoverArrow: {
    // Fluent autoSize writes inline overflow-x/y:auto when constraining the
    // surface. Important class rules keep its outside arrow visible; the inner
    // content wrapper owns scrolling. Kebab-case keys use Griffel's CSS rule API.
    'overflow-x': 'visible !important',
    'overflow-y': 'visible !important',
  },
  popoverSmall: { width: '240px' },
  popoverMedium: { width: '320px' },
  popoverLarge: { width: '480px' },
});

type StageStyles = ReturnType<typeof useSpecificStyles> & {
  content: string;
  modalPadding: string;
  popoverPadding: string;
  popoverContent: string;
};

/** Shared host-rendered classes. Adapters combine surface + presentation + size + padding.
 * Only the surface owns padding: modalPadding is 24px, popoverPadding is 12px.
 * content is a zero-padding flex column for an adapter-owned content wrapper, if needed.
 * dialog/drawer/popover carry viewport limits independently from the size classes.
 */
export function useStageStyles(): StageStyles {
  const styles = useSpecificStyles(), sx = useSx();
  return {
    ...styles,
    surface: sx(styles.surface, flex.column, minHeight.zero, minWidth.zero, gap.medium, overflow.auto),
    content: sx(flex.column, flexItem.grow, minHeight.zero, minWidth.zero, gap.medium, padding.none),
    // Scroll content inside the native surface so its outside anchor arrow is not clipped.
    popoverContent: sx(flex.column, flexItem.grow, minHeight.zero, minWidth.zero, gap.medium, padding.none, overflow.auto),
    header: sx(styles.header, flexItem.noShrink, minWidth.zero, padding.none),
    body: sx(styles.body, flexItem.grow, minHeight.zero, minWidth.zero, overflow.auto, padding.none),
    footer: sx(styles.footer, flexItem.noShrink, minWidth.zero, padding.none),
    neutralCard: sx(styles.neutralCard, padding.none, gap.none, borderRadius.none, width.full, container.inlineSize,
      background.transparent, boxShadow.none, overflow.visible, minWidth.zero),
    // Native image/action margins provide horizontal separation only for present slots.
    cardHeader: sx(styles.cardHeader, gap.none, padding.none),
    text: sx(styles.text, flex.column, gap.none, minWidth.zero),
    description: sx(styles.description, minWidth.zero),
    icon: sx(styles.icon, flexItem.noShrink),
    actions: sx(styles.actions, flex.row, flex.wrap, alignItems.center, gap.small, flexItem.noShrink, minWidth.zero),
    modalPadding: sx(padding.px(24)),
    popoverPadding: sx(padding.px(12)),
  };
}
