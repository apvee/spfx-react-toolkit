import {
  width, minWidth, height, minHeight, maxWidth, maxHeight, flex, flexItem, grid,
  alignItems, justifyContent, alignSelf, gap, padding, paddingInline, paddingBlock,
  paddingInlineStart, paddingInlineEnd, paddingBlockStart, paddingBlockEnd,
  margin, marginInline, marginBlock, marginInlineStart, marginInlineEnd,
  marginBlockStart, marginBlockEnd, foreground, background, presets, typography,
  textAlign, text, borderWidth, borderStyle, borderColor, borderRadius, boxShadow,
  overflow, scrollbar, container, SxBaseDescriptor,
} from '@apvee/spfx-react-toolkit';

type Selection = SxBaseDescriptor | ((value: number) => SxBaseDescriptor);
interface CatalogFamily {
  readonly name: string;
  readonly members: Readonly<Record<string, Selection>>;
}
const family = (name: string, members: Readonly<Record<string, Selection>>): CatalogFamily => ({ name, members });

/** Private interactive choices; these are real public descriptors, never CSS substitutes. */
export const stylesCatalog: readonly CatalogFamily[] = [
  family('width', width), family('minWidth', minWidth), family('height', height),
  family('minHeight', minHeight), family('maxWidth', maxWidth), family('maxHeight', maxHeight),
  family('flex', flex), family('flexItem', flexItem), family('grid', grid),
  family('alignItems', alignItems), family('justifyContent', justifyContent), family('alignSelf', alignSelf),
  family('gap', gap), family('padding', padding), family('paddingInline', paddingInline),
  family('paddingBlock', paddingBlock), family('paddingInlineStart', paddingInlineStart),
  family('paddingInlineEnd', paddingInlineEnd), family('paddingBlockStart', paddingBlockStart),
  family('paddingBlockEnd', paddingBlockEnd), family('margin', margin), family('marginInline', marginInline),
  family('marginBlock', marginBlock), family('marginInlineStart', marginInlineStart),
  family('marginInlineEnd', marginInlineEnd), family('marginBlockStart', marginBlockStart),
  family('marginBlockEnd', marginBlockEnd), family('foreground', foreground), family('background', background),
  family('presets', presets), family('typography', typography), family('textAlign', textAlign), family('text', text),
  family('borderWidth', borderWidth), family('borderStyle', borderStyle), family('borderColor', borderColor),
  family('borderRadius', borderRadius), family('boxShadow', boxShadow),
  family('overflow', { visible: overflow.visible, hidden: overflow.hidden, auto: overflow.auto }),
  family('overflow.horizontal', overflow.horizontal), family('overflow.vertical', overflow.vertical),
  family('scrollbar', scrollbar), family('container', container),
];
