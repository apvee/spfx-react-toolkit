import type { TitleColor, TitleDocument, TitleHeadingLevel, TitleLineSpacing } from './web-part-title.types.internal';
import { webLightTheme, type PartialTheme } from '@fluentui/react-theme';

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
    brand: 'colorBrandForeground1'
  },
  spaceToken: 'spacingVerticalS',
  legacySizes: { fontSizeXLarge: 20 },
  legacyColors: { fontColorBlue: 'brand' }
} as const;

export function isTitleSize(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) &&
    value >= titleFormatRegistry.minSize && value <= titleFormatRegistry.maxSize;
}

export function isTitleColor(value: unknown): value is TitleColor {
  return value === 'primary' || value === 'subtle' || value === 'brand';
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
