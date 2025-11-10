// theme-subscription.ts
// Utility hook to subscribe to SPFx theme changes

import { useEffect } from 'react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ThemeProvider } from '@microsoft/sp-component-base';
import type { ServiceScope } from '@microsoft/sp-core-library';

/**
 * Extract ThemeProvider from SPFx context
 * Works with any SPFx context type (WebPart, ApplicationCustomizer, etc.)
 * @internal
 */
function getThemeProvider(spfxContext: unknown): ThemeProvider {
  const ctx = spfxContext as {
    serviceScope?: ServiceScope;
  };
  
  if (!ctx.serviceScope) {
    throw new Error('SPFx context does not have serviceScope');
  }
  
  // Consume ThemeProvider from service scope
  // ServiceScope.consume() is type-safe with ServiceKey<T>
  return ctx.serviceScope.consume(ThemeProvider.serviceKey);
}

/**
 * Hook to subscribe to SPFx theme changes
 * Automatically updates the provided setter when theme changes
 * 
 * @param spfxContext - SPFx context object
 * @param setTheme - Setter function to update theme state
 * @internal
 */
export function useThemeSubscription(
  spfxContext: unknown,
  setTheme: (theme: IReadonlyTheme | undefined) => void
): void {
  useEffect(() => {
    const themeProvider = getThemeProvider(spfxContext);
    
    // Get initial theme
    const initialTheme = themeProvider.tryGetTheme();
    if (initialTheme) {
      setTheme(initialTheme);
    }
    
    // Create event handler
    const handler = (args: { theme?: IReadonlyTheme }): void => {
      setTheme(args.theme);
    };
    
    // Create observer object for SPFx event system
    const observer = {
      instanceId: 'theme-subscription',
      componentId: 'theme-subscription',
      isDisposed: false,
      dispose: (): void => {
        // Cleanup handled in useEffect return
      },
      update: handler,
    };
    
    // Subscribe to theme changes
    themeProvider.themeChangedEvent.add(observer, handler);
    
    // Cleanup on unmount
    return () => {
      observer.isDisposed = true;
      themeProvider.themeChangedEvent.remove(observer, handler);
    };
  }, [spfxContext, setTheme]);
}
