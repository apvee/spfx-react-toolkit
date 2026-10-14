import type { TitleColor, TitleDocument, TitleHeadingLevel, TitleLineSpacing } from './web-part-title.types.internal';
import { tokens, webLightTheme, type PartialTheme } from '@fluentui/react-theme';

type TitleForegroundToken = Extract<keyof typeof tokens, `color${string}Foreground${string}`>;
type TitleForegroundSource = TitleForegroundToken | {
  readonly first: TitleForegroundToken;
  readonly second: TitleForegroundToken;
  readonly percent: 25 | 40 | 50 | 55 | 60 | 73 | 75;
};

/** Immutable protocol registry shared by codec, renderer and editor. No runtime registrations. */
export const titleFormatRegistry = {
  htmlByteLimit: 64 * 1024,
  depthLimit: 32,
  minSize: 8,
  maxSize: 96,
  suggestedSizes: [8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72, 96],
  headingPresets: {
    1: { typography: 'title1', sizeToken: 'fontSizeHero800' },
    2: { typography: 'title2', sizeToken: 'fontSizeHero700' },
    3: { typography: 'title3', sizeToken: 'fontSizeBase600' },
    4: { typography: 'subtitle1', sizeToken: 'fontSizeBase500' }
  },
  paragraphPreset: { typography: 'body1', sizeToken: 'fontSizeBase300' },
  lineSpacing: [1, 1.2, 1.4, 1.6, 2, 2.4],
  colors: {
    primary: 'colorNeutralForeground1',
    subtle: 'colorNeutralForeground2',
    brand: 'colorBrandForeground1',
    neutralMuted: 'colorNeutralForeground3',
    brandAlt: 'colorBrandForeground2',
    red: 'colorPaletteRedForeground3',
    orange: 'colorPaletteDarkOrangeForeground3',
    yellow: 'colorPaletteYellowForeground3',
    green: 'colorPaletteGreenForeground3',
    teal: 'colorPaletteTealForeground2',
    blue: 'colorPaletteBlueForeground2',
    purple: 'colorPalettePurpleForeground2',
    pink: 'colorPalettePinkForeground2',
    brown: 'colorPaletteBrownForeground2',
    neutralDark: { first: 'colorNeutralForeground1', second: 'colorNeutralForeground2', percent: 50 },
    neutralLight: { first: 'colorNeutralForeground3', second: 'colorNeutralForegroundInverted', percent: 75 },
    neutralLighter: { first: 'colorNeutralForeground3', second: 'colorNeutralForegroundInverted', percent: 60 },
    brandDarkest: { first: 'colorBrandForeground1', second: 'colorNeutralForeground1', percent: 25 },
    brandDarker: { first: 'colorBrandForeground1', second: 'colorNeutralForeground1', percent: 40 },
    brandDark: { first: 'colorBrandForeground1', second: 'colorNeutralForeground1', percent: 55 },
    brandLight: { first: 'colorBrandForeground1', second: 'colorNeutralForegroundInverted', percent: 75 },
    darkRed: 'colorPaletteRedForeground1',
    amber: 'colorPaletteMarigoldForeground3',
    yellowGreen: { first: 'colorPaletteYellowForeground3', second: 'colorPaletteGreenForeground3', percent: 73 },
    cyan: 'colorPaletteLightTealForeground2',
    navy: 'colorPaletteNavyForeground2'
  },
  spaceToken: 'spacingVerticalS',
  legacySizes: { fontSizeXLarge: 20 },
  legacyColors: { fontColorBlue: 'brand' }
} as const;

/** Ordered private palette. Legacy IDs remain valid and appear only while selected. */
export const titleColorDescriptors = [
  { id: 'primary', group: 'theme' },
  { id: 'neutralDark', group: 'theme' },
  { id: 'subtle', group: 'theme' },
  { id: 'neutralMuted', group: 'theme' },
  { id: 'neutralLight', group: 'theme' },
  { id: 'neutralLighter', group: 'theme' },
  { id: 'brandDarkest', group: 'theme' },
  { id: 'brandDarker', group: 'theme' },
  { id: 'brandDark', group: 'theme' },
  { id: 'brandAlt', group: 'theme' },
  { id: 'brand', group: 'theme' },
  { id: 'brandLight', group: 'theme' },
  { id: 'darkRed', group: 'palette' },
  { id: 'red', group: 'palette' },
  { id: 'amber', group: 'palette' },
  { id: 'yellow', group: 'palette' },
  { id: 'yellowGreen', group: 'palette' },
  { id: 'green', group: 'palette' },
  { id: 'cyan', group: 'palette' },
  { id: 'blue', group: 'palette' },
  { id: 'navy', group: 'palette' },
  { id: 'purple', group: 'palette' },
  { id: 'orange', group: 'palette', legacy: true },
  { id: 'teal', group: 'palette', legacy: true },
  { id: 'pink', group: 'palette', legacy: true },
  { id: 'brown', group: 'palette', legacy: true }
] as const satisfies readonly { id: TitleColor; group: 'theme' | 'palette'; legacy?: true }[];

/** Shared foreground-only CSS resolver for swatches and scoped Read/Edit classes. */
export function getTitleColorCss(color: TitleColor): string {
  const source: TitleForegroundSource = titleFormatRegistry.colors[color];
  return typeof source === 'string' ? tokens[source]
    : `color-mix(in srgb, ${tokens[source.first]} ${source.percent}%, ${tokens[source.second]})`;
}

/** Resolve standard Fluent hex values for deduplication; custom CSS stays symbolic.
 * CSS variables and system colors cannot be resolved without a browser context.
 * Equal custom endpoints still collapse, while distinct expressions remain available.
 */
export function getTitleColorValueKey(color: TitleColor, theme?: PartialTheme): string {
  const source: TitleForegroundSource = titleFormatRegistry.colors[color];
  const value = (token: TitleForegroundToken): string => {
    const resolved = (theme?.[token] ?? tokens[token]).trim();
    // Custom property names are case-sensitive; preserve CSS expressions verbatim.
    if (/^#[\da-f]{3}$/i.test(resolved)) return '#' + resolved.slice(1).toLowerCase().split('').map(character => character + character).join('');
    return /^#[\da-f]{6}$/i.test(resolved) || /^[a-z]+$/i.test(resolved) ? resolved.toLowerCase() : resolved;
  };
  if (typeof source === 'string') return value(source);
  const first = value(source.first);
  const second = value(source.second);
  if (first === second) return first;
  const channels = (input: string): readonly number[] | undefined => {
    const match = /^#([\da-f]{6}|[\da-f]{3})$/.exec(input);
    if (!match) return undefined;
    const hex = match[1].length === 3 ? match[1].split('').map(character => character + character).join('') : match[1];
    return [0, 2, 4].map(offset => Number.parseInt(hex.slice(offset, offset + 2), 16));
  };
  const a = channels(first);
  const b = channels(second);
  if (a && b) {
    return '#' + a.map((channel, index) => Math.round((channel * source.percent + b[index] * (100 - source.percent)) / 100)
      .toString(16).padStart(2, '0')).join('');
  }
  return `color-mix(in srgb, ${first} ${source.percent}%, ${second})`;
}

/** Deduplicate offered colors while retaining the active semantic ID and Automatic. */
export function getTitlePickerColors(selected: TitleColor | 'mixed' | undefined, theme?: PartialTheme): typeof titleColorDescriptors[number][] {
  const activeValue = selected !== undefined && selected !== 'mixed' ? getTitleColorValueKey(selected, theme) : undefined;
  const seen = new Set<string>();
  return titleColorDescriptors.filter(item => {
    if ('legacy' in item && item.id !== selected) return false;
    const value = getTitleColorValueKey(item.id, theme);
    if (item.id === selected) { seen.add(value); return true; }
    if (value === activeValue || seen.has(value)) return false;
    seen.add(value);
    return true;
  });
}

export function isTitleSize(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) &&
    value >= titleFormatRegistry.minSize && value <= titleFormatRegistry.maxSize;
}

export function isTitleColor(value: unknown): value is TitleColor {
  return typeof value === 'string' && Object.prototype.hasOwnProperty.call(titleFormatRegistry.colors, value);
}

export function isTitleHeadingLevel(value: unknown): value is TitleHeadingLevel {
  return value === 1 || value === 2 || value === 3 || value === 4;
}

export function isTitleLineSpacing(value: unknown): value is TitleLineSpacing {
  return value === 1 || value === 1.2 || value === 1.4 || value === 1.6 || value === 2 || value === 2.4;
}

/** Resolve the preset's pixel size from the active Fluent theme, without persisting it. */
export function getTitleDefaultSize(document: { type: TitleDocument['type']; level?: TitleHeadingLevel }, theme?: PartialTheme): number {
  const preset = document.type === 'heading' && document.level
    ? titleFormatRegistry.headingPresets[document.level] : titleFormatRegistry.paragraphPreset;
  const size = Number.parseFloat(theme?.[preset.sizeToken] ?? webLightTheme[preset.sizeToken]);
  return Number.isFinite(size) ? size : Number.parseFloat(webLightTheme[preset.sizeToken]);
}

export function getTitleSizeClass(size: number): string | undefined {
  return isTitleSize(size) ? `wps-size-${size}` : undefined;
}

export function getTitleLineClass(line: TitleLineSpacing): string {
  return `wps-line-${Math.round(line * 10)}`;
}

export function getTitleClassSize(className: string): number | undefined {
  if (className === 'fontSizeXLarge') return titleFormatRegistry.legacySizes.fontSizeXLarge;
  const match = /^wps-size-([1-9]\d?)$/.exec(className);
  const size = match ? Number(match[1]) : undefined;
  return isTitleSize(size) ? size : undefined;
}

export function getTitleClassColor(className: string): TitleColor | undefined {
  if (className === 'fontColorBlue') return titleFormatRegistry.legacyColors.fontColorBlue;
  const value = className.startsWith('wps-color-') ? className.slice(10) : undefined;
  return isTitleColor(value) ? value : undefined;
}

export function getTitleClassLine(className: string): TitleLineSpacing | undefined {
  if (!/^wps-line-(10|12|14|16|20|24)$/.test(className)) return undefined;
  const value = Number(className.slice(9)) / 10;
  return isTitleLineSpacing(value) ? value : undefined;
}
