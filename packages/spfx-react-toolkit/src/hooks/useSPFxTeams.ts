// useSPFxTeams.ts
// Hook for Microsoft Teams context integration

import { useEffect, useMemo } from 'react';
import { useSPFxRuntimeStore, useSPFxRuntimeSelector } from '../core/state.internal';
import { useSPFxContext } from './useSPFxContext';
import { resolveSPFxTeamsSdk } from '../helpers/spfx-host-context.internal';
import { subscribeSPFxTeamsRuntime } from '../core/teams-runtime.internal';

/**
 * Teams theme type
 */
export type TeamsTheme = 'default' | 'dark' | 'highContrast';

/**
 * Return type for useSPFxTeams hook
 */
export interface SPFxTeamsInfo {
  /** Whether context acquisition from the SPFx application SDK succeeded. */
  readonly supported: boolean;
  
  /** Teams context object (if available) */
  readonly context: unknown | undefined;
  
  /** Current Teams theme */
  readonly theme: TeamsTheme | undefined;
}

/**
 * Hook for Microsoft Teams context integration
 * 
 * Provides access to Microsoft Teams context when SPFx component
 * exposes the Microsoft 365 application SDK, including Teams, Office and Outlook.
 * 
 * Acquires context from the SPFx-provided SDK (supports both v1 and v2 APIs)
 * and provides:
 * - Teams context (team, channel, user info, etc.)
 * - Teams theme (default, dark, high contrast)
 * - Supported flag
 * 
 * Context acquisition happens asynchronously on first mount. Live theme events
 * remain subscribed until the last hook for this runtime store unmounts. SDK
 * replacement resets the context and ignores obsolete completions. Registration
 * and acquisition failures are reported through console.error.
 *
 * @remarks
 * A lazy broker shares one theme dispatcher per SDK identity within this library
 * bundle. The SDK has no public unregister API; cleanup releases runtime stores.
 * Other bundles that register their own SDK handler can replace the dispatcher.
 * The supported flag means SDK context is available; useSPFxEnvironmentInfo
 * resolves whether the actual host is Teams, Office or Outlook.
 * 
 * Useful for:
 * - Teams-specific features
 * - Theme synchronization
 * - User context access
 * - Channel/team information
 * - Teams app integration
 * 
 * @returns Teams context information
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { supported, context, theme } = useSPFxTeams();
 *   
 *   if (!supported) {
 *     return <div>Not in Teams</div>;
 *   }
 *   
 *   const teamsContext = context as {
 *     user?: { id: string };
 *     team?: { displayName: string };
 *   };
 *   
 *   return (
 *     <div className={`teams-theme-${theme}`}>
 *       <p>Team: {teamsContext.team?.displayName}</p>
 *       <p>Theme: {theme}</p>
 *     </div>
 *   );
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
