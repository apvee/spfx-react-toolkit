// useSPFxTeams.ts
// Hook for Microsoft Teams context integration

import { useEffect } from 'react';
import { useAtom } from 'jotai';
import { spfxAtoms } from '../core/atoms';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';
import { useSPFxContext } from './useSPFxContext';

/**
 * Teams theme type
 */
export type TeamsTheme = 'default' | 'dark' | 'highContrast';

/**
 * Return type for useSPFxTeams hook
 */
export interface SPFxTeamsInfo {
  /** Whether Teams context is supported/available */
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
 * is running in Teams environment.
 * 
 * Automatically initializes Teams SDK (supports both v1 and v2 APIs)
 * and provides:
 * - Teams context (team, channel, user info, etc.)
 * - Teams theme (default, dark, high contrast)
 * - Supported flag
 * 
 * The initialization happens asynchronously on first mount.
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
  const { id } = useSPFxInstanceInfo();
  const { spfxContext } = useSPFxContext();
  const [state, setState] = useAtom(spfxAtoms.teams(id));
  
  useEffect(() => {
    // Skip if already initialized
    if (state.initialized) {
      return;
    }
    
    // Extract Teams SDK from context
    const ctx = spfxContext as { sdks?: { microsoftTeams?: unknown } };
    const sdks = ctx.sdks;
    const teamsSDK = sdks?.microsoftTeams as
      | {
          teamsJs?: unknown;
          app?: { getContext?: () => Promise<unknown> };
          getContext?: (callback: (context: unknown) => void) => void;
        }
      | undefined;
    
    if (!teamsSDK) {
      setState({ supported: false, initialized: true });
      return;
    }
    
    let disposed = false;
    
    // Helper to apply context and theme
    const apply = (context: unknown, themeValue?: string): void => {
      if (disposed) {
        return;
      }
      
      const normalizedTheme = (
        themeValue?.toLowerCase() === 'dark' ? 'dark' :
        themeValue?.toLowerCase() === 'contrast' ||  themeValue?.toLowerCase() === 'highcontrast' ? 'highContrast' :
        'default'
      ) as TeamsTheme;
      
      setState({
        supported: true,
        context,
        theme: normalizedTheme,
        initialized: true,
      });
    };
    
    // Try v2 API (teams 2.0+)
    const tryV2 = async (): Promise<boolean> => {
      try {
        if (teamsSDK.app?.getContext) {
          const context = await teamsSDK.app.getContext();
          const themeValue = (context as { app?: { theme?: string } })?.app?.theme;
          apply(context, themeValue);
          return true;
        }
        return false;
      } catch {
        return false;
      }
    };
    
    // Try v1 API (teams 1.x)
    const tryV1 = (): Promise<boolean> => {
      return new Promise<boolean>((resolve) => {
        try {
          if (teamsSDK.getContext) {
            teamsSDK.getContext((context: unknown) => {
              const themeValue = (context as { theme?: string })?.theme;
              apply(context, themeValue);
              resolve(true);
            });
          } else {
            resolve(false);
          }
        } catch {
          resolve(false);
        }
      });
    };
    
    // Initialize async
    const init = async (): Promise<void> => {
      const v2Success = await tryV2();
      if (!v2Success) {
        await tryV1();
      }
    };
    
    init().catch(() => {
      // Fallback: mark as not supported
      if (!disposed) {
        setState({ supported: false, initialized: true });
      }
    });
    
    return () => {
      disposed = true;
    };
  }, [id, spfxContext, setState, state.initialized]);
  
  return {
    supported: state.supported,
    context: state.context,
    theme: state.theme,
  };
}
