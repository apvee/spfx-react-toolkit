/**
 * Application SDK context and theme subscriptions belong to the provider runtime.
 *
 * @packageDocumentation
 * @docPage hooks/environment
 */
// useSPFxTeams.ts
// Hook for Microsoft Teams context integration

import { useEffect, useMemo } from 'react';
import { useSPFxRuntimeStore, useSPFxRuntimeSelector } from '../core/state.internal';
import { useSPFxContext } from './useSPFxContext';
import { resolveSPFxTeamsSdk } from '../helpers/spfx-host-context.internal';
import { subscribeSPFxTeamsRuntime } from '../core/teams-runtime.internal';

/**
 * Teams theme type
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { TeamsTheme } from '@apvee/spfx-react-toolkit';
 *
 * const theme: TeamsTheme = 'highContrast';
 * ```
 */
export type TeamsTheme = 'default' | 'dark' | 'highContrast';

/**
 * Asynchronously acquired application SDK context and normalized theme; supported is not a Teams-host discriminator.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxTeamsInfo, useSPFxTeams } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const host: SPFxTeamsInfo = useSPFxTeams();
 *   console.log(host.supported, host.theme);
 *   return <span>{host.theme ?? 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxTeamsInfo {
  /** Whether context acquisition from the SPFx application SDK succeeded. */
  readonly supported: boolean;
  
  /**
   * Unverified host SDK context; undefined before acquisition or when unavailable. Narrow using the SDK contract and actual application host.
   */
  readonly context: unknown | undefined;
  
  /**
   * Normalized application theme, or undefined before/unavailable SDK acquisition.
   */
  readonly theme: TeamsTheme | undefined;
}

/**
 * Acquires SPFx-provided application SDK context asynchronously, supporting v1/v2 SDK shapes and Teams, Office and Outlook hosts. supported is false until successful acquisition and is not a Teams-host discriminator; use useSPFxEnvironmentInfo for that. Missing SDK/context leaves context and theme undefined. The last subscriber releases the runtime store; SDK replacement resets context and ignores obsolete acquisitions. A dispatcher is shared per SDK identity within this bundle; another bundle can replace its single theme handler. Failures are logged to console.error. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns SDK readiness, untyped host context and optional normalized application theme.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read Teams metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTeams } from '@apvee/spfx-react-toolkit';
 *
 * function ApplicationTheme() {
 *   const { supported, theme } = useSPFxTeams();
 *   return <span>{supported ? theme ?? 'Theme unavailable' : 'SDK context unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen a component needs the acquired application SDK context and its published runtime theme.
 *
 * @example Observe application SDK readiness and theme
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxTeams } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxTeams();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxTeams result">
 *         <dt>SDK context acquired</dt><dd>{String(info.supported)}</dd>
 *         <dt>SDK context available</dt><dd>{String(info.context !== undefined)}</dd>
 *         <dt>SDK theme</dt><dd>{info.theme ?? 'Unavailable'}</dd>
 *       </dl>,
 *     diagnostics: <p>Supported means SDK context acquisition succeeded, including Office or Outlook. It is not a Teams host discriminator. The native SDK owns theme callbacks; cleanup releases provider subscribers.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual Teams values and explicit unavailable fields as the native host changes.' };
 * ```
 */
export function useSPFxTeams(): SPFxTeamsInfo {
  const { spfxContext } = useSPFxContext();
  const state = useSPFxRuntimeSelector(runtimeState => runtimeState.teams);
  const store = useSPFxRuntimeStore();
  const sdk = resolveSPFxTeamsSdk(spfxContext);

  useEffect(() => {
    if (!sdk) {
      store.setState({ teams: { supported: false, initialized: true } });
      return;
    }
    return subscribeSPFxTeamsRuntime(sdk, store);
  }, [sdk, store]);

  return useMemo(() => ({
    supported: state.supported,
    context: state.context,
    theme: state.theme,
  }), [state.supported, state.context, state.theme]);
}
