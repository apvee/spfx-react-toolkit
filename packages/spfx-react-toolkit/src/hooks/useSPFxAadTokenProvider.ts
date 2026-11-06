/**
 * Native token-provider acquisition does not request tokens automatically.
 *
 * @packageDocumentation
 * @docPage hooks/http-clients/aad-token-provider
 */
// useSPFxAadTokenProvider.ts
// Hook to access the SPFx Azure AD token provider with state management

import { AadTokenProviderFactory } from '@microsoft/sp-http';
import type { AadTokenProvider } from '@microsoft/sp-http';
import { useEffect, useMemo, useRef, useState } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';

/**
 * Native SPFx token provider with asynchronous initialization state; acquisition failures remain distinct from token-request failures.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxAadTokenProviderInfo, useSPFxAadTokenProvider } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const info: SPFxAadTokenProviderInfo = useSPFxAadTokenProvider();
 *   console.log(info.isReady, info.initError);
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
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
 * Acquires the SPFx AadTokenProvider asynchronously from ServiceScope. Factory consumption and provider acquisition failures are captured in initError. No token is acquired by this hook itself.
 * @returns Provider, initialization state and readiness.
 * @throws When called outside an SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxAadTokenProvider } from '@apvee/spfx-react-toolkit';
 *
 * function TokenStatus() {
 *   const { isReady, initError } = useSPFxAadTokenProvider();
 *   return <span>{initError?.message ?? (isReady ? 'Ready' : 'Initializing')}</span>;
 * }
 * ```
 *
 * @useWhen an integration needs the native token provider rather than an HTTP wrapper.
 *
 *
 * @example Inspect delegated scopes from a manual token
 * ```tsx
 * import * as React from 'react';
 * import {
 *   decodeSPFxJwtPayload,
 *   useSPFxAadTokenProvider,
 * } from '@apvee/spfx-react-toolkit';
 *
 * function TokenProbe() {
 *   const { tokenProvider, isReady, isInitializing, initError } = useSPFxAadTokenProvider();
 *   const [scopeText, setScopeText] = React.useState<string>('');
 *   const [failure, setFailure] = React.useState<string>();
 *
 *   const inspectGraphToken = async () => {
 *     if (!tokenProvider) {
 *       return;
 *     }
 *
 *     setFailure(undefined);
 *     try {
 *       const token = await tokenProvider.getToken('https://graph.microsoft.com');
 *       const payload = decodeSPFxJwtPayload(token);
 *       setScopeText(payload?.scp || '');
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *
 *   if (isInitializing) return <p role="status">Preparing token provider...</p>;
 *   if (initError) return <p role="alert">{initError.message}</p>;
 *
 *   return (
 *     <>
 *       <button onClick={inspectGraphToken} disabled={!isReady}>
 *         Inspect Graph scopes
 *       </button>
 *       {failure && <p role="alert">{failure}</p>}
 *       <pre>{scopeText}</pre>
 *     </>
 *   );
 * }
 * ```
 *
 * @scenario Inspect client readiness and token-provider initialization
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHttpClient, useSPFxSPHttpClient, useSPFxMSGraphClient, useSPFxAadHttpClient, useSPFxAadTokenProvider, useSPFxServiceScope } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const { serviceScope } = useSPFxServiceScope();
 *   const http = useSPFxHttpClient();
 *   const sp = useSPFxSPHttpClient();
 *   const graph = useSPFxMSGraphClient();
 *   const aad = useSPFxAadHttpClient();
 *   const token = useSPFxAadTokenProvider();
 *   // The hooks clear old initialization in passive effects. Mask the first new-scope commit.
 *   const [observedScope, setObservedScope] = React.useState(serviceScope);
 *   React.useEffect(() => { setObservedScope(serviceScope); }, [serviceScope]);
 *   const current = observedScope === serviceScope;
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="Client readiness">
 *       <dt>HttpClient</dt><dd>{http.isReady ? 'Ready' : 'Unavailable'}</dd>
 *       <dt>SPHttpClient</dt><dd>{sp.isReady ? 'Ready' : 'Unavailable'}</dd>
 *       <dt>MSGraphClient</dt><dd>{current && graph.isReady ? 'Ready' : 'Unavailable'}</dd>
 *       <dt>AadHttpClient</dt><dd>{current && aad.isReady ? 'Ready' : 'Unavailable'}</dd>
 *       <dt>Token provider</dt><dd>Token provider: {current && token.isReady ? 'Ready' : 'Unavailable'}</dd>
 *       <dt>SPHttp Base URL</dt><dd>{sp.baseUrl}</dd>
 *       <dt>Graph Initializing</dt><dd>{!current || graph.isInitializing ? 'Yes' : 'No'}</dd>
 *       <dt>AAD Resource</dt><dd>{aad.resourceUrl || 'Not configured'}</dd>
 *       <dt>AAD Initializing</dt><dd>{!current || aad.isInitializing ? 'Yes' : 'No'}</dd>
 *       <dt>AAD Token Provider Initializing</dt><dd>{!current || token.isInitializing ? 'Yes' : 'No'}</dd>
 *     </dl>,
 *     status: <>{current && [graph.initError, aad.initError, token.initError].map((error, index) => error && <p role="alert" key={index}>{error.message}</p>)}</>,
 *     diagnostics: <p>Mount below the matching SPFx provider. Factories initialize clients/providers without business requests. No token is requested. Reset leaves native initialization and readiness intact.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Watch actual client and token-provider initialization, readiness and initialization failures without requesting a token.' };
 * ```
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
