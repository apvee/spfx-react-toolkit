import { tokens, webDarkTheme, webLightTheme } from '@fluentui/react-theme';
import type { Theme } from '@fluentui/react-theme';
import { resolveThemeMode } from '../../helpers/fluent-theme-mode.internal';

function isThemeToken(name: string): name is keyof Theme {
  return Object.prototype.hasOwnProperty.call(tokens, name);
}
function tokenValue(style: CSSStyleDeclaration, name: keyof Theme): string | undefined {
  const value = style.getPropertyValue(`--${name}`).trim();
  return value && !/^(initial|inherit|unset|revert|revert-layer)$/i.test(value) && !value.includes('var(') ? value : undefined;
}
// The existing mode resolver accepts hex/comma-RGB but intentionally defaults
// unknown colors to light. Validate that bounded input first so an unusable
// transient snapshot cannot replace the previous complete native theme.
function isOpaqueColor(value: string): boolean {
  if (/^#(?:[\da-f]{3}|[\da-f]{6})$/i.test(value)) return true;
  if (/^#[\da-f]{4}$/i.test(value)) return /f$/i.test(value);
  if (/^#[\da-f]{8}$/i.test(value)) return /ff$/i.test(value);
  const match = /^rgba?\(([^)]+)\)$/i.exec(value);
  if (!match) return false;
  const parts = match[1].split(',').map(part => part.trim());
  if (parts.length !== 3 && parts.length !== 4) return false;
  return parts.every((part, index) => {
    if (!/^(?:\d+(?:\.\d*)?|\.\d+)%?$/.test(part)) return false;
    const maximum = part.endsWith('%') ? 100 : index === 3 ? 1 : 255;
    const number = Number.parseFloat(part);
    return Number.isFinite(number) && (index === 3 ? number === maximum : number >= 0 && number <= maximum);
  });
}
function copyToken<K extends keyof Theme>(theme: Theme, name: K, value: string): void {
  const numeric = typeof webLightTheme[name] === 'number';
  const parsed = numeric ? Number(value) : value;
  if (numeric && (!Number.isFinite(parsed) || Number(parsed) < 1 || Number(parsed) > 1000)) return;
  // The canonical token key determines its installed Theme type: four numeric
  // font weights; all remaining tokens (including optional z-index) are strings.
  theme[name] = parsed as Theme[K];
}
/** Read the native custom-field container, never its themed React descendants. */
export function readPropertyPanePickerTheme(container: HTMLElement): Theme | undefined {
  const owner = container.ownerDocument.defaultView;
  if (!container.isConnected || !owner) return undefined;
  const style = owner.getComputedStyle(container);
  const background = tokenValue(style, 'colorNeutralBackground1');
  const foreground = tokenValue(style, 'colorNeutralForeground1');
  if (background === undefined || foreground === undefined || !isOpaqueColor(background) || !isOpaqueColor(foreground)) return undefined;
  const base = resolveThemeMode({ colorNeutralBackground1: background, colorNeutralForeground1: foreground }) === 'dark' ? webDarkTheme : webLightTheme;
  const theme = { ...base };
  for (const name of Object.keys(tokens)) {
    if (!isThemeToken(name)) continue;
    const value = tokenValue(style, name);
    if (value !== undefined) copyToken(theme, name, value);
  }
  return theme;
}
/** Equivalent native snapshots retain ThemeContext identity and do not remount consumers. */
export function samePropertyPanePickerTheme(first: Theme, second: Theme): boolean {
  return Object.keys(tokens).every(name => !isThemeToken(name) || first[name] === second[name]);
}
