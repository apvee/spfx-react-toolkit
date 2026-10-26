import type { EditableTextColor } from '../editable-text-document.types.internal';

/** Control-owned resources. Author content and persisted formatting identifiers are never translated. */
export interface EditableTextStrings {
  readonly editorLoadError: string;
  readonly retry: string;
  readonly savedFormattingWarning: string;
  readonly loadingEditor: string;
  readonly editorLabel: string;
  readonly placeholder: string;
  readonly tooLargeError: string;
  readonly moreFormatting: string;
  readonly heading: string;
  readonly normal: string;
  readonly bold: string;
  readonly italic: string;
  readonly underline: string;
  readonly themeColor: string;
  readonly clearFormatting: string;
  readonly spacing: string;
  readonly defaultLineSpacing: string;
  readonly spaceBefore: string;
  readonly spaceAfter: string;
  readonly formattingControls: string;
  readonly formatting: string;
  readonly fontSize: string;
  readonly mixed: string;
  readonly themeColorGroup: string;
  readonly automatic: string;
  readonly themeColors: string;
  readonly paletteColors: string;
  readonly colorNames: Readonly<Record<EditableTextColor, string>>;
  readonly headingLevel: (level: number) => string;
  readonly lineSpacingValue: (value: number) => string;
  readonly formattingShortcut: (label: string, shortcut: string) => string;
  readonly mixedFormatting: (label: string) => string;
}

type FormattedKeys = 'headingLevel' | 'lineSpacingValue' | 'formattingShortcut' | 'mixedFormatting';
/** Templates keep grammatical word order in each catalog; command shortcuts remain explicit inputs. */
export type EditableTextStringResource = Omit<EditableTextStrings, FormattedKeys> & {
  readonly headingLevelTemplate: `${string}{level}${string}`;
  readonly lineSpacingValueTemplate: `${string}{value}${string}`;
  readonly formattingShortcutTemplate:
    | `${string}{label}${string}{shortcut}${string}`
    | `${string}{shortcut}${string}{label}${string}`;
  readonly mixedFormattingTemplate: `${string}{label}${string}`;
};

/** Construct an immutable, instance-independent catalog from a fully typed resource. */
export function defineEditableTextStrings(resource: EditableTextStringResource): EditableTextStrings {
  const { headingLevelTemplate, lineSpacingValueTemplate, formattingShortcutTemplate, mixedFormattingTemplate, ...labels } = resource;
  return Object.freeze({
    ...labels,
    colorNames: Object.freeze({ ...resource.colorNames }),
    headingLevel: (level: number): string => headingLevelTemplate.replace('{level}', String(level)),
    lineSpacingValue: (value: number): string => lineSpacingValueTemplate.replace('{value}', String(value)),
    formattingShortcut: (label: string, shortcut: string): string => formattingShortcutTemplate
      .replace(/\{(label|shortcut)\}/g, (_match, token: string) => token === 'label' ? label : shortcut),
    mixedFormatting: (label: string): string => mixedFormattingTemplate.replace('{label}', () => label),
  });
}
