import type { TitleColor } from '../title/web-part-title.types.internal';

/** Control-owned resources. Author content and persisted formatting identifiers are never translated. */
export interface WebPartShellStrings {
  readonly titleEditorLoadError: string;
  readonly retry: string;
  readonly savedTitleFormattingWarning: string;
  readonly loadingTitleEditor: string;
  readonly titleEditorLabel: string;
  readonly titlePlaceholder: string;
  readonly titleTooLargeError: string;
  readonly moreTitleFormatting: string;
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
  readonly titleFormattingControls: string;
  readonly titleFormatting: string;
  readonly fontSize: string;
  readonly mixed: string;
  readonly titleThemeColors: string;
  readonly automatic: string;
  readonly themeColors: string;
  readonly paletteColors: string;
  readonly colorNames: Readonly<Record<TitleColor, string>>;
  readonly headingLevel: (level: number) => string;
  readonly lineSpacingValue: (value: number) => string;
  readonly formattingShortcut: (label: string, shortcut: string) => string;
  readonly mixedFormatting: (label: string) => string;
}

type FormattedKeys = 'headingLevel' | 'lineSpacingValue' | 'formattingShortcut' | 'mixedFormatting';
/** Templates keep grammatical word order in each catalog; command shortcuts remain explicit inputs. */
export type WebPartShellStringResource = Omit<WebPartShellStrings, FormattedKeys> & {
  readonly headingLevelTemplate: `${string}{level}${string}`;
  readonly lineSpacingValueTemplate: `${string}{value}${string}`;
  readonly formattingShortcutTemplate:
    | `${string}{label}${string}{shortcut}${string}`
    | `${string}{shortcut}${string}{label}${string}`;
  readonly mixedFormattingTemplate: `${string}{label}${string}`;
};

/** Construct an immutable, instance-independent catalog from a fully typed resource. */
export function defineWebPartShellStrings(resource: WebPartShellStringResource): WebPartShellStrings {
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
