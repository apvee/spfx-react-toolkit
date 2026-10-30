import type { Theme } from '@fluentui/react-theme';

export type ThemeMode = 'light' | 'dark';
type RgbColor = readonly [red: number, green: number, blue: number];

function parseHexColor(value: string): RgbColor | undefined {
  const match = /^#([\da-f]{3,4}|[\da-f]{6}|[\da-f]{8})$/i.exec(value);
  if (!match) return undefined;
  const hex = match[1];
  const short = hex.length <= 4;
  const channel = (offset: number): number => Number.parseInt(
    short ? hex.slice(offset, offset + 1).repeat(2) : hex.slice(offset * 2, offset * 2 + 2), 16
  );
  // The alpha component is intentionally ignored, matching the source artwork theme heuristic.
  return [channel(0), channel(1), channel(2)];
}

function parseNumber(value: string, maximum: number): number | undefined {
  if (!/^(?:\d+(?:\.\d*)?|\.\d+)%?$/.test(value)) return undefined;
  const percentage = value.endsWith('%');
  const parsed = Number.parseFloat(value);
  const limit = percentage ? 100 : maximum;
  if (!Number.isFinite(parsed) || parsed > limit) return undefined;
  return percentage ? (parsed / 100) * maximum : parsed;
}

function parseRgbFunction(value: string): RgbColor | undefined {
  const match = /^rgba?\(\s*([^,\s]+)\s*,\s*([^,\s]+)\s*,\s*([^,\s)]+)\s*(?:,\s*([^,\s)]+)\s*)?\)$/i.exec(value);
  if (!match) return undefined;
  if (match[4] !== undefined && parseNumber(match[4], 1) === undefined) return undefined;
  const red = parseNumber(match[1], 255);
  const green = parseNumber(match[2], 255);
  const blue = parseNumber(match[3], 255);
  if (red === undefined || green === undefined || blue === undefined) return undefined;
  return [red, green, blue];
}

function parseColor(value: string | undefined): RgbColor | undefined {
  if (value === undefined) return undefined;
  const normalized = value.trim();
  return parseHexColor(normalized) ?? parseRgbFunction(normalized);
}

function linearize(channel: number): number {
  const normalized = channel / 255;
  return normalized <= 0.04045 ? normalized / 12.92 : ((normalized + 0.055) / 1.055) ** 2.4;
}

function relativeLuminance([red, green, blue]: RgbColor): number {
  return 0.2126 * linearize(red) + 0.7152 * linearize(green) + 0.0722 * linearize(blue);
}

/** Heuristic based on surface/text luminance; absent or unsupported theme colors use light artwork. */
export function resolveThemeMode(theme: Partial<Theme> | undefined): ThemeMode {
  const background = parseColor(theme?.colorNeutralBackground1);
  const foreground = parseColor(theme?.colorNeutralForeground1);
  if (background === undefined || foreground === undefined) return 'light';
  return relativeLuminance(background) < relativeLuminance(foreground) ? 'dark' : 'light';
}
