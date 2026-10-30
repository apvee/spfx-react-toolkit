import type { EditableTextStrings, EditableTextStringsOverride } from '../editable-text.types';
import type { EditableTextResource } from './editable-text-resource.types.internal';
import { nativeEditableTextResource } from './editable-text-native.internal';
import { formatResourceTemplate } from '../../../localization/resource-formatters.internal';

/** Compose formatter callbacks once from native data without leaking translator metadata. */
export function defineEditableTextStrings(resource: EditableTextResource): EditableTextStrings {
  const { headingLevelTemplate, lineSpacingValueTemplate, formattingShortcutTemplate, mixedFormattingTemplate } = resource;
  // Select only public observable fields: native metadata is loader-owned.
  return Object.freeze({
    editorLoadError: resource.editorLoadError,
    retry: resource.retry,
    savedFormattingWarning: resource.savedFormattingWarning,
    loadingEditor: resource.loadingEditor,
    placeholder: resource.placeholder,
    tooLargeError: resource.tooLargeError,
    moreFormatting: resource.moreFormatting,
    heading: resource.heading,
    normal: resource.normal,
    bold: resource.bold,
    italic: resource.italic,
    underline: resource.underline,
    themeColor: resource.themeColor,
    clearFormatting: resource.clearFormatting,
    spacing: resource.spacing,
    defaultLineSpacing: resource.defaultLineSpacing,
    spaceBefore: resource.spaceBefore,
    spaceAfter: resource.spaceAfter,
    formattingControls: resource.formattingControls,
    formatting: resource.formatting,
    fontSize: resource.fontSize,
    mixed: resource.mixed,
    themeColorGroup: resource.themeColorGroup,
    automatic: resource.automatic,
    themeColors: resource.themeColors,
    paletteColors: resource.paletteColors,
    colorNames: Object.freeze({ ...resource.colorNames }),
    headingLevel: (level: number): string => formatResourceTemplate(headingLevelTemplate, { level }),
    lineSpacingValue: (value: number): string => formatResourceTemplate(lineSpacingValueTemplate, { value }),
    formattingShortcut: (label: string, shortcut: string): string => formatResourceTemplate(formattingShortcutTemplate, { label, shortcut }),
    mixedFormatting: (label: string): string => formatResourceTemplate(mixedFormattingTemplate, { label }),
  });
}

const nativeStrings = defineEditableTextStrings(nativeEditableTextResource);

/** Defined own overrides win; nested color names inherit individually. */
export function resolveEditableTextStrings(overrides?: EditableTextStringsOverride): EditableTextStrings {
  if (!overrides) return nativeStrings;
  const defined = Object.fromEntries(Object.entries(overrides).filter(([key, value]) => key !== 'colorNames' && value !== undefined));
  const colors = Object.fromEntries(Object.entries(overrides.colorNames ?? {}).filter(([, value]) => value !== undefined));
  return { ...nativeStrings, ...defined, colorNames: { ...nativeStrings.colorNames, ...colors } };
}
