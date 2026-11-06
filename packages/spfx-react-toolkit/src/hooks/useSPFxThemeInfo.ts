/**
 * Native SPFx themes update through the provider theme subscription.
 *
 * @packageDocumentation
 * @docPage hooks/theming
 */
// useSPFxThemeInfo.ts
// Hook to access current SPFx theme

import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { useSPFxRuntimeSelector } from '../core/state.internal';

/**
 * Reads the nearest provider's host SPFx theme reference. Provider subscriptions publish changes; absence remains undefined, without a local theme fallback. Use useSPFxFluent9ThemeInfo to obtain Fluent UI 9 theme tokens. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Host readonly theme, or undefined when unavailable.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read ThemeInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxThemeInfo } from '@apvee/spfx-react-toolkit';
 *
 * function ThemeLabel() {
 *   const theme = useSPFxThemeInfo();
 *   return <span>{theme ? 'Host theme available' : 'Host theme unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen a component needs the current native SPFx theme and theme-change updates.
 *
 * @scenario Observe the native SPFx theme snapshot
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxThemeInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxThemeInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxThemeInfo result">
 *         <dt>Native theme</dt><dd>{info ? 'Available' : 'Unavailable'}</dd>
 *         <dt>Inverted</dt><dd>{info ? String(info.isInverted) : 'Unavailable'}</dd>
 *         <dt>Theme primary</dt><dd>{info?.palette?.themePrimary ?? 'Unavailable'}</dd>
 *         <dt>Body text</dt><dd>{info?.semanticColors?.bodyText ?? 'Unavailable'}</dd>
 *       </dl>,
 *     diagnostics: <p>Native ThemeProvider subscriptions publish updates. Absence stays undefined. This example inherits the host presentation and does not manufacture an SPFx theme.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual ThemeInfo values and explicit unavailable fields as the native host changes.' };
 * ```
 */
export function useSPFxThemeInfo(): IReadonlyTheme | undefined {
  return useSPFxRuntimeSelector(state => state.theme);
}
