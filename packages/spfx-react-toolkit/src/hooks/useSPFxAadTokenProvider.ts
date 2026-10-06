// useSPFxAadTokenProvider.ts
// Hook to access the SPFx Azure AD token provider with state management

import { AadTokenProviderFactory } from '@microsoft/sp-http';
import type { AadTokenProvider } from '@microsoft/sp-http';
import { useEffect, useMemo, useRef, useState } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';

/**
 * Return type for useSPFxAadTokenProvider hook
 */
export interface SPFxAadTokenProviderInfo {
  /**
   * Native AadTokenProvider from SPFx.
   * Provides access tokens for Azure AD-secured resources.
   * Will be undefined until initialization completes.
   */
  readonly tokenProvider: AadTokenProvider | undefined;

  /**
   * True while the AAD token provider is being initialized.
   * Use this to show a loading indicator during startup.
   */
  readonly isInitializing: boolean;

  /**
   * Error that occurred during token provider initialization.
   * If set, the token provider will remain undefined.
   */
  readonly initError: Error | undefined;

  /**
   * Computed state: true when token provider is ready for use.
   * Equivalent to: tokenProvider !== undefined && !isInitializing && !initError
   */
  readonly isReady: boolean;
}

/**
 * Hook to access SPFx AadTokenProvider with built-in state management.
 *
 * @remarks
 * This hook consumes AadTokenProviderFactory from SPFx ServiceScope using
 * dependency injection. The factory is consumed lazily and cached. The
 * factory.getTokenProvider() method is then called asynchronously to obtain
 * the AadTokenProvider instance.
 */
export function useSPFxAadTokenProvider(): SPFxAadTokenProviderInfo {
  const { consume } = useSPFxServiceScope();

  // ═══════════════════════════════════════════════════════════════════════════
  // STATE
  // ═══════════════════════════════════════════════════════════════════════════

  const [tokenProvider, setTokenProvider] = useState<AadTokenProvider | undefined>(undefined);
  const [isInitializing, setIsInitializing] = useState<boolean>(true);
  const [initError, setInitError] = useState<Error | undefined>(undefined);

  // ═══════════════════════════════════════════════════════════════════════════
  // REFS (for cleanup and preventing stale updates)
  // ═══════════════════════════════════════════════════════════════════════════

  // Track component mounted state to prevent memory leaks
  const isMountedRef = useRef<boolean>(true);
  const requestIdRef = useRef<number>(0);

  // Cleanup on unmount
  useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  // ═══════════════════════════════════════════════════════════════════════════
  // FACTORY (lazy consume from ServiceScope)
  // ═══════════════════════════════════════════════════════════════════════════

  // Lazy consume AadTokenProviderFactory from ServiceScope (cached by useMemo)
  const factoryResult = useMemo(() => {
    try {
      return {
        factory: consume<AadTokenProviderFactory>(AadTokenProviderFactory.serviceKey),
        error: undefined,
      };
    } catch (err: unknown) {
      return {
        factory: undefined,
        error: err instanceof Error ? err : new Error(String(err)),
      };
    }
  }, [consume]);

  // ═══════════════════════════════════════════════════════════════════════════
  // INITIALIZATION EFFECT
  // ═══════════════════════════════════════════════════════════════════════════

  // Initialize AAD token provider (factory.getTokenProvider is async)
  useEffect(() => {
    const requestId = requestIdRef.current + 1;
    requestIdRef.current = requestId;

    // Reset state for new initialization
    setTokenProvider(undefined);
    setIsInitializing(true);
    setInitError(undefined);

    if (factoryResult.error || !factoryResult.factory) {
      if (isMountedRef.current && requestId === requestIdRef.current) {
        setInitError(factoryResult.error);
        setIsInitializing(false);
        console.error('Failed to consume AadTokenProviderFactory:', factoryResult.error);
      }
      return;
    }

    // Get AadTokenProvider
    factoryResult.factory
      .getTokenProvider()
      .then((aadTokenProvider: AadTokenProvider) => {
        // Only update state if still mounted and this request is current
        if (isMountedRef.current && requestId === requestIdRef.current) {
          setTokenProvider(aadTokenProvider);
          setIsInitializing(false);
        }
      })
      .catch((err: unknown) => {
        // Only update state if still mounted and this request is current
        if (isMountedRef.current && requestId === requestIdRef.current) {
          const error = err instanceof Error ? err : new Error(String(err));
          setInitError(error);
          setIsInitializing(false);
          console.error('Failed to initialize AadTokenProvider:', error);
        }
      });
  }, [factoryResult]);

  // ═══════════════════════════════════════════════════════════════════════════
  // COMPUTED STATE & RETURN
  // ═══════════════════════════════════════════════════════════════════════════

  // Computed: ready when token provider is available and no errors
  const isReady = tokenProvider !== undefined && !isInitializing && !initError;

  return useMemo(() => ({
    tokenProvider,
    isInitializing,
    initError,
    isReady,
  }), [tokenProvider, isInitializing, initError, isReady]);
}
