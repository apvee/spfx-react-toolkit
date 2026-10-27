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
 * const theme: TeamsTheme = 'highContrast';
 * ```
 */
export type TeamsTheme = 'default' | 'dark' | 'highContrast';

/**
 * Return type for useSPFxTeams hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const host: SPFxTeamsInfo = useSPFxTeams();
 * console.log(host.supported, host.theme);
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
 * @example
 * ```tsx
 * function ApplicationTheme() {
 *   const { supported, theme } = useSPFxTeams();
 *   return <span>{supported ? theme ?? 'Theme unavailable' : 'SDK context unavailable'}</span>;
 * }
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
