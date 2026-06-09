// useSPFxApiPermissionPrecheck.ts
// React facade for SPFx API permission diagnostics

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { ISPEventObserver } from '@microsoft/sp-core-library';
import type { AadTokenProvider, BeforeRedirectEventArgs, PopupEventArgs } from '@microsoft/sp-http';
import {
  createSPFxApiPermissionPrecheckService,
} from '../services/spfx-api-permission-precheck.service';
import {
  normalizeSPFxApiPermissionRequirements,
  summarizeSPFxApiPermissionResults,
  type SPFxApiPermissionCheckResult,
  type SPFxApiPermissionConfigurationState,
  type SPFxApiPermissionPrecheckConfig,
  type SPFxApiPermissionPrecheckMode,
  type SPFxApiPermissionRequirement,
  type SPFxApiPermissionSummary,
} from '../helpers/spfx-api-permission-precheck.helpers';
import { useSPFxAadTokenProvider } from './useSPFxAadTokenProvider';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';

export interface SPFxApiPermissionPrecheckOptions {
  readonly autoCheck?: boolean;
  readonly mode?: SPFxApiPermissionPrecheckMode;
  readonly useCachedToken?: boolean;
  readonly validateAudience?: boolean;
  readonly timeoutMs?: number;
}

export interface SPFxApiPermissionPrecheckResult {
  readonly isChecking: boolean;
  readonly isConfigured: boolean;
  readonly configurationState: SPFxApiPermissionConfigurationState;
  readonly available: readonly SPFxApiPermissionSummary[];
  readonly missing: readonly SPFxApiPermissionSummary[];
  readonly warnings: readonly SPFxApiPermissionSummary[];
  readonly unknown: readonly SPFxApiPermissionSummary[];
  readonly results: readonly SPFxApiPermissionCheckResult[];
  readonly tokenProviderError: Error | undefined;
  readonly check: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  readonly retry: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  readonly retryWithoutCache: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
}

interface SPFxEventLike<TEventArgs> {
  add(observer: ISPEventObserver, eventHandler: (eventArgs: TEventArgs) => void): void;
  remove(observer: ISPEventObserver, eventHandler: (eventArgs: TEventArgs) => void): void;
}

interface AadTokenProviderPassiveEvents {
  readonly popupEvent?: SPFxEventLike<PopupEventArgs>;
  readonly onBeforeRedirectEvent?: SPFxEventLike<BeforeRedirectEventArgs>;
}

interface PassiveAuthEventObserverInfo {
  readonly instanceId: string;
  readonly componentId: string;
}

export function useSPFxApiPermissionPrecheck(
  config: SPFxApiPermissionPrecheckConfig,
  options?: SPFxApiPermissionPrecheckOptions
): SPFxApiPermissionPrecheckResult {
  const {
    tokenProvider,
    isInitializing,
    initError,
    isReady
  } = useSPFxAadTokenProvider();
  const instanceInfo = useSPFxInstanceInfo();

  const autoCheck = options?.autoCheck !== false;
  const mode = options?.mode || 'passive';
  const useCachedToken = options?.useCachedToken !== false;
  const validateAudience = options?.validateAudience !== false;
  const timeoutMs = options?.timeoutMs ?? 15000;

  const normalizedRequirements = useMemo(
    () => normalizeSPFxApiPermissionRequirements(config),
    [config]
  );
  const service = useMemo(
    () => tokenProvider ? createSPFxApiPermissionPrecheckService(tokenProvider) : undefined,
    [tokenProvider]
  );
  const autoCheckKey = useMemo(
    () => createAutoCheckKey(config, {
      autoCheck,
      mode,
      useCachedToken,
      validateAudience,
      timeoutMs
    }),
    [autoCheck, config, mode, timeoutMs, useCachedToken, validateAudience]
  );

  const [isChecking, setIsChecking] = useState<boolean>(false);
  const [results, setResults] = useState<readonly SPFxApiPermissionCheckResult[]>([]);
  const isMountedRef = useRef<boolean>(true);
  const requestIdRef = useRef<number>(0);
  const autoCheckKeyRef = useRef<string | undefined>(undefined);

  useEffect(() => {
    isMountedRef.current = true;

    return () => {
      isMountedRef.current = false;
    };
  }, []);

  useEffect(() => {
    if (tokenProvider || isInitializing) {
      return;
    }

    const requestId = requestIdRef.current + 1;
    requestIdRef.current = requestId;
    const unavailableResults = createTokenProviderUnavailableResults(normalizedRequirements);

    if (isMountedRef.current && requestId === requestIdRef.current) {
      setIsChecking(false);
      setResults(unavailableResults);
    }
  }, [isInitializing, normalizedRequirements, tokenProvider]);

  const runCheck = useCallback(async (
    nextUseCachedToken: boolean
  ): Promise<readonly SPFxApiPermissionCheckResult[]> => {
    const requestId = requestIdRef.current + 1;
    requestIdRef.current = requestId;

    if (isMountedRef.current) {
      setIsChecking(true);
    }

    try {
      const nextResults = service && tokenProvider
        ? await runWithPassiveAuthEventGuard(
          tokenProvider,
          mode,
          {
            instanceId: instanceInfo.id,
            componentId: instanceInfo.kind
          },
          () => service.check(config, {
            useCachedToken: nextUseCachedToken,
            validateAudience,
            timeoutMs
          })
        )
        : createTokenProviderUnavailableResults(normalizedRequirements);

      if (isMountedRef.current && requestId === requestIdRef.current) {
        setResults(nextResults);
      }

      return nextResults;
    } finally {
      if (isMountedRef.current && requestId === requestIdRef.current) {
        setIsChecking(false);
      }
    }
  }, [
    config,
    instanceInfo.id,
    instanceInfo.kind,
    mode,
    normalizedRequirements,
    service,
    timeoutMs,
    tokenProvider,
    validateAudience
  ]);

  const check = useCallback(
    () => runCheck(useCachedToken),
    [runCheck, useCachedToken]
  );

  const retry = useCallback(
    () => check(),
    [check]
  );

  const retryWithoutCache = useCallback(
    () => runCheck(false),
    [runCheck]
  );

  useEffect(() => {
    if (!autoCheck || !isReady || !tokenProvider) {
      return;
    }

    if (autoCheckKeyRef.current === autoCheckKey) {
      return;
    }

    autoCheckKeyRef.current = autoCheckKey;
    check().catch(() => undefined);
  }, [autoCheck, autoCheckKey, check, isReady, tokenProvider]);

  const summary = useMemo(
    () => summarizeSPFxApiPermissionResults(results, isChecking),
    [isChecking, results]
  );

  return useMemo(() => ({
    isChecking,
    isConfigured: summary.isConfigured,
    configurationState: summary.configurationState,
    available: summary.available,
    missing: summary.missing,
    warnings: summary.warnings,
    unknown: summary.unknown,
    results,
    tokenProviderError: initError,
    check,
    retry,
    retryWithoutCache
  }), [
    check,
    initError,
    isChecking,
    results,
    retry,
    retryWithoutCache,
    summary.available,
    summary.configurationState,
    summary.isConfigured,
    summary.missing,
    summary.unknown,
    summary.warnings
  ]);
}

async function runWithPassiveAuthEventGuard(
  tokenProvider: AadTokenProvider,
  mode: SPFxApiPermissionPrecheckMode,
  observerInfo: PassiveAuthEventObserverInfo,
  run: () => Promise<readonly SPFxApiPermissionCheckResult[]>
): Promise<readonly SPFxApiPermissionCheckResult[]> {
  if (mode !== 'passive') {
    return run();
  }

  const providerWithEvents = tokenProvider as unknown as AadTokenProviderPassiveEvents;
  const popupEvent = providerWithEvents.popupEvent;
  const redirectEvent = providerWithEvents.onBeforeRedirectEvent;
  const popupHandler = (eventArgs: PopupEventArgs): void => {
    eventArgs.cancel(new Error('SPFx API permission precheck passive mode blocked an authentication popup.'));
  };
  const redirectHandler = (eventArgs: BeforeRedirectEventArgs): void => {
    eventArgs.cancel();
  };
  const observer = createPassiveAuthEventObserver(observerInfo);
  let popupRegistered = false;
  let redirectRegistered = false;

  try {
    if (hasEventRegistrationApi(popupEvent)) {
      popupEvent.add(observer, popupHandler);
      popupRegistered = true;
    }

    if (hasEventRegistrationApi(redirectEvent)) {
      redirectEvent.add(observer, redirectHandler);
      redirectRegistered = true;
    }

    return await run();
  } finally {
    if (popupRegistered && hasEventRegistrationApi(popupEvent)) {
      popupEvent.remove(observer, popupHandler);
    }

    if (redirectRegistered && hasEventRegistrationApi(redirectEvent)) {
      redirectEvent.remove(observer, redirectHandler);
    }

    observer.dispose();
  }
}

function createPassiveAuthEventObserver(
  observerInfo: PassiveAuthEventObserverInfo
): ISPEventObserver {
  const mutableObserver = {
    instanceId: observerInfo.instanceId || 'spfx-api-permission-precheck',
    componentId: observerInfo.componentId || 'spfx-api-permission-precheck',
    isDisposed: false,
    dispose: (): void => {
      mutableObserver.isDisposed = true;
    }
  };

  return mutableObserver;
}

function hasEventRegistrationApi<TEventArgs>(
  event: SPFxEventLike<TEventArgs> | undefined
): event is SPFxEventLike<TEventArgs> {
  return typeof event?.add === 'function' && typeof event.remove === 'function';
}

function createTokenProviderUnavailableResults(
  requirements: readonly SPFxApiPermissionRequirement[]
): readonly SPFxApiPermissionCheckResult[] {
  return requirements.map(requirement => ({
    id: requirement.id,
    resourceName: requirement.resourceName,
    resourceEndpoint: requirement.resourceEndpoint,
    scope: requirement.scope,
    required: requirement.required,
    status: 'tokenProviderUnavailable',
    severity: requirement.required ? 'error' : 'warning',
    message: `The SPFx AadTokenProvider is unavailable, so ${requirement.resourceName} / ${requirement.scope} cannot be checked.`,
    adminMessage: requirement.adminMessage,
    packageSolutionEntry: {
      resource: requirement.packageResource,
      scope: requirement.scope
    },
    detectedScopes: []
  }));
}

function createAutoCheckKey(
  config: SPFxApiPermissionPrecheckConfig,
  options: Required<SPFxApiPermissionPrecheckOptions>
): string {
  try {
    return JSON.stringify({ config, options });
  } catch {
    return 'unserializable-config';
  }
}
