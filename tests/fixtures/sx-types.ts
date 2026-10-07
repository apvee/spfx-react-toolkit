import {
  width, minWidth, foreground, typography, hover, active, focusVisible, responsive, viewport,
  SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor, SxInput, SxFunction, SxOptions
} from '../../packages/spfx-react-toolkit/src/helpers/styles';

const base: SxBaseDescriptor = width.full;
const state: SxStateDescriptor = hover(foreground.subtle);
const query: SxResponsiveDescriptor = responsive.medium(hover(foreground.subtle));
const inputs: readonly SxInput[] = [base, state, query, width.px(240), minWidth.zero, typography.body1,
  active(width.full), focusVisible(typography.body1), viewport.small(width.auto), 'existing-class', false, null, undefined];
const options: SxOptions = { dir: 'rtl' };
declare const sx: SxFunction;
const className: string = sx(...inputs);
void options;
void className;
// @ts-expect-error Raw CSS objects cannot satisfy an opaque descriptor.
const raw: SxBaseDescriptor = { width: '100%' };
// @ts-expect-error Pixel factories accept numbers only.
width.px('240px');
// @ts-expect-error Scopes do not accept compiled class strings.
hover('className');
// @ts-expect-error Responsive scopes do not accept class strings.
responsive.medium('className');
// @ts-expect-error Queries cannot be nested.
responsive.medium(viewport.small(width.full));
// @ts-expect-error States cannot be nested.
hover(active(width.full));
// @ts-expect-error Responsive descriptors cannot be wrapped in a state.
focusVisible(responsive.medium(width.full));
// @ts-expect-error CSS objects are not composer inputs.
sx({ width: '240px' });
// @ts-expect-error Opaque descriptor members are immutable.
base.extra = true;
void raw;

import {
  height, minHeight, maxWidth, maxHeight, flex, flexItem, grid, alignItems, justifyContent, alignSelf,
  gap, padding, paddingInline, paddingBlock, paddingInlineStart, paddingInlineEnd, paddingBlockStart, paddingBlockEnd,
  margin, marginInline, marginBlock, marginInlineStart, marginInlineEnd, marginBlockStart, marginBlockEnd,
  background, presets, textAlign, text
} from '../../packages/spfx-react-toolkit/src/helpers/styles';
const catalog: readonly SxBaseDescriptor[] = [height.full, minHeight.zero, maxWidth.none, maxHeight.px(200),
  flex.row, flexItem.grow, grid.columns(2), alignItems.center, justifyContent.spaceBetween, alignSelf.auto,
  gap.medium, padding.small, paddingInline.extraSmall, paddingBlock.large, paddingInlineStart.none,
  paddingInlineEnd.extraLarge, paddingBlockStart.px(12), paddingBlockEnd.medium,
  margin.small, marginInline.medium, marginBlock.large, marginInlineStart.none, marginInlineEnd.px(2),
  marginBlockStart.extraSmall, marginBlockEnd.extraLarge, background.alternative, presets.brandTint,
  textAlign.end, text.truncate, typography.display, foreground.link];
const catalogClass: string = sx(...catalog);
void catalogClass;
// @ts-expect-error Grid columns require a numeric count.
grid.columns('2');
// @ts-expect-error Spacing pixel factories accept only numbers.
padding.px('2px');

import { borderWidth, borderStyle, borderColor, borderRadius, boxShadow } from '../../packages/spfx-react-toolkit/src/helpers/styles';
const decorated: SxBaseDescriptor[] = [borderWidth.thin, borderStyle.solid, borderColor.brand, borderRadius.circular, boxShadow.small];
sx(...decorated);

import { overflow, scrollbar } from '../../packages/spfx-react-toolkit/src/helpers/styles';
sx(overflow.auto, overflow.horizontal.hidden, overflow.vertical.auto, scrollbar.fluent);
// @ts-expect-error The v1 scrollbar catalog has no additional recipes.
scrollbar.webKit;
