/**
 * Delegated-scope diagnostics do not grant access or replace server authorization.
 *
 * @packageDocumentation
 * @docPage hooks/http-clients/api-permission-precheck
 */
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

/**
 * Controls automatic checks, authentication interaction and token evaluation.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxApiPermissionPrecheckOptions, useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const options: SPFxApiPermissionPrecheckOptions = { autoCheck: false, mode: 'passive', timeoutMs: 10000 };
 *   const permissions = useSPFxApiPermissionPrecheck({ graph: ['User.Read'] }, options);
 *   return <span>{permissions.configurationState}</span>;
 * }
 * ```
 */
export interface SPFxApiPermissionPrecheckOptions {
  /**
   * Automatically checks when the token provider is ready and requirements/options change; omitted or undefined uses true.
   * @defaultValue true
   */
  readonly autoCheck?: boolean;
  /**
   * Omitted or undefined uses passive: resource acquisition is sequential and available popup/redirect events are cancelled for the check. interactiveAllowed does not install that guard; the provider controls actual interaction.
   * @defaultValue 'passive'
   */
  readonly mode?: SPFxApiPermissionPrecheckMode;
  /**
   * Allows cached tokens unless false; omitted or undefined uses true.
   * @defaultValue true
   */
  readonly useCachedToken?: boolean;
  /**
   * Checks expected audiences unless false; omitted or undefined uses true.
   * @defaultValue true
   */
  readonly validateAudience?: boolean;
  /**
   * Per-resource token acquisition timeout in milliseconds; omitted or undefined uses 15000. Does not cancel the underlying provider request.
   * @defaultValue 15000
   */
  readonly timeoutMs?: number;
}

/**
 * Permission diagnostics and retry operations. The latest-started check publishes state; obsolete/unmounted completions still settle their promises but cannot publish results. Diagnostics do not grant access.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxApiPermissionPrecheckResult, useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const permissions: SPFxApiPermissionPrecheckResult = useSPFxApiPermissionPrecheck({ graph: ['User.Read'] });
 *   console.log(permissions.configurationState, permissions.missing);
 *   return <span>{permissions.configurationState}</span>;
 * }
 * ```
 */
export interface SPFxApiPermissionPrecheckResult {
  /**
   * True while the latest check is pending.
   */
  readonly isChecking: boolean;
  /**
   * True when every required evaluated result is available. Empty/no-required results also yield true, even while configurationState is idle/checking; does not establish server authorization.
   */
  readonly isConfigured: boolean;
  /**
   * Aggregate diagnostic state, including idle, checking and actionRequired.
   */
  readonly configurationState: SPFxApiPermissionConfigurationState;
  /**
   * Requirements evaluated as available.
   */
  readonly available: readonly SPFxApiPermissionSummary[];
  /**
   * Required unavailable requirements with actionable missing/blocked access diagnostics.
   */
  readonly missing: readonly SPFxApiPermissionSummary[];
  /**
   * All unavailable optional requirements; they do not block isConfigured.
   */
  readonly warnings: readonly SPFxApiPermissionSummary[];
  /**
   * Required unavailable requirements without an actionable missing-permission diagnosis.
   */
  readonly unknown: readonly SPFxApiPermissionSummary[];
  /**
   * Detailed per-scope diagnostics; no raw tokens are included.
   */
  readonly results: readonly SPFxApiPermissionCheckResult[];
  /**
   * Token-provider initialization error, independent of per-scope check diagnostics.
   */
  readonly tokenProviderError: Error | undefined;
  /**
   * Runs a check with the configured cache preference.
   * @returns Per-scope diagnostics, including unavailable-provider results.
   * @throws Unexpected normalization or authentication-event registration failures.
   */
  readonly check: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  /**
   * Repeats check with the configured cache preference.
   * @returns Fresh diagnostics.
   * @throws Unexpected check failures.
   */
  readonly retry: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  /**
   * Repeats check requesting uncached tokens; the passive/interactive mode still applies.
   * @returns Fresh diagnostics.
   * @throws Unexpected check failures.
   */
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

/**
 * Checks delegated scopes using the provider's AadTokenProvider. Auto-check defaults to true and mode to passive. Requirement/provider changes reset results; latest checks publish state, without cancelling underlying requests. Token failures become diagnostics.
 * @param config - Graph/custom API requirements; keep data immutable when changing requirements.
 * @param options - Automation/auth settings; omitted or undefined uses auto-check, passive mode, cached tokens, audience validation and 15000 ms timeout.
 * @returns Aggregate state, detailed diagnostics and explicit retry methods.
 * @throws When called outside an SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';
 *
 * function PermissionStatus() {
 *   const permissions = useSPFxApiPermissionPrecheck({ graph: ['User.Read'] });
 *   return <span>{permissions.configurationState}</span>;
 * }
 * ```
 *
 * @useWhen setup UI needs delegated-scope diagnostics and actionable retry results.
 *
 * @avoidWhen the application needs proof that a server request will succeed; token diagnostics do not grant access.
 *
 *
 * @example Compare a gate and manual passive recheck
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';
 * function OrdersPermissionGate({ children }: { children: React.ReactNode }) {
 *   const precheck = useSPFxApiPermissionPrecheck({ graph: ['Sites.Read.All'], customApis: [{
 *     name: 'Orders API', resource: 'api://contoso-orders-api', packageResource: 'Orders API', scopes: ['Orders.Read'],
 *   }] });
 *   return precheck.configurationState === 'ready' ? <>{children}</> : <section>
 *     <p>{precheck.isChecking ? 'Checking API access' : precheck.configurationState}</p>
 *     {precheck.missing.map(item => <p key={item.id}>{item.adminMessage}</p>)}
 *   </section>;
 * }
 * function ManualPermissionCheck() {
 *   const precheck = useSPFxApiPermissionPrecheck({ graph: ['Sites.Read.All'] }, { autoCheck: false, mode: 'passive' });
 *   return <div><button disabled={precheck.isChecking} onClick={() => void precheck.check().catch(console.error)}>Check now</button>
 *     <button disabled={precheck.isChecking} onClick={() => void precheck.retryWithoutCache().catch(console.error)}>Recheck without cache</button>
 *     <pre>{precheck.configurationState}</pre>
 *   </div>;
 * }
 * ```
 *
 * @scenario Manually check Graph and custom API delegated scopes
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Accordion, AccordionItem, AccordionHeader, AccordionPanel, makeStyles } from '@fluentui/react-components';
 * import { useSPFxApiPermissionPrecheck, useSPFxAadTokenProvider, useSPFxContext, useSPFxServiceScope, summarizeSPFxApiPermissionResults, SPFxApiPermissionCheckResult } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ diagnostics: { minWidth: 0, maxWidth: '100%', overflowX: 'auto' } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const styles = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const { serviceScope } = useSPFxServiceScope();
 *   const token = useSPFxAadTokenProvider();
 *   const [resource, setResource] = React.useState('');
 *   const target = resource.trim();
 *   const config = React.useMemo(() => ({ graph: ['User.Read'], customApis: target ? [{ name: 'Custom API', resource: target, scopes: ['user_impersonation'] }] : [] }), [target]);
 *   const options = React.useMemo(() => ({ autoCheck: false, mode: 'passive' } as const), []);
 *   const info = useSPFxApiPermissionPrecheck(config, options);
 *   const [observedScope, setObservedScope] = React.useState(serviceScope);
 *   React.useEffect(() => { setObservedScope(serviceScope); }, [serviceScope]);
 *   const values: readonly unknown[] = [spfxContext, serviceScope, serviceScope?.consume, target, resetRevision, info.check, info.retryWithoutCache];
 *   // A transition gets a new boundary even if a later draft/revision repeats an old value.
 *   const boundary = React.useRef(values);
 *   if (!values.every((value, index) => boundary.current[index] === value)) boundary.current = values;
 *   const identity = boundary.current;
 *   const live = React.useRef(identity);
 *   live.current = identity;
 *   const mounted = React.useRef(true);
 *   React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const same = (capture: readonly unknown[]): boolean => mounted.current && capture === live.current && capture[1] === spfxContext.serviceScope && capture[2] === serviceScope?.consume;
 *   const [completed, setCompleted] = React.useState<{ identity: readonly unknown[]; results?: readonly SPFxApiPermissionCheckResult[]; failure?: string }>();
 *   // A configuration change may clear hook isChecking before its original promise settles.
 *   const originalPending = React.useRef(false);
 *   const [pending, setPending] = React.useState(false);
 *   const releaseOriginal = (): void => {
 *     originalPending.current = false;
 *     if (mounted.current) setPending(false);
 *   };
 *   const run = async (withoutCache: boolean): Promise<void> => {
 *     if (!same(identity) || originalPending.current || observedScope !== serviceScope || !token.isReady) return;
 *     originalPending.current = true;
 *     setPending(true);
 *     setCompleted(undefined);
 *     try {
 *       const results = await (withoutCache ? info.retryWithoutCache() : info.check());
 *       if (same(identity)) setCompleted({ identity, results });
 *     } catch (failure) {
 *       if (same(identity)) setCompleted({ identity, failure: failure instanceof Error ? failure.message : String(failure) });
 *     } finally {
 *       // Settlement releases this original promise's latch even when its view became obsolete.
 *       releaseOriginal();
 *     }
 *   };
 *   const baseline = React.useRef(resetRevision);
 *   React.useEffect(() => { if (baseline.current !== resetRevision) { baseline.current = resetRevision; setResource(''); setCompleted(undefined); } }, [resetRevision]);
 *   const visible = completed && same(completed.identity) ? completed : undefined;
 *   const summary = summarizeSPFxApiPermissionResults(visible?.results ?? [], pending || info.isChecking);
 *   const disabled = observedScope !== serviceScope || !token.isReady || pending || info.isChecking;
 *   const regions: DemoRegions = {
 *     controls: <Field label="Resource URL or App ID URI"><Input aria-label="Resource URL or App ID URI" value={resource} onChange={(_, data) => setResource(data.value)} /></Field>,
 *     commands: <><Button disabled={disabled} onClick={() => run(false)}>Check permissions</Button><Button disabled={disabled} onClick={() => run(true)}>Retry without cache</Button></>,
 *     status: <>{pending && <p role="status">Previous check may still be completing. Reset does not cancel native work.</p>}{visible?.failure && <p role="alert">{visible.failure}</p>}{observedScope === serviceScope && info.tokenProviderError && <p role="alert">{info.tokenProviderError.message}</p>}</>,
 *     children: <><dl aria-label="API permission precheck results"><dt>Configuration State</dt><dd>{summary.configurationState}</dd><dt>Configured</dt><dd>{summary.isConfigured ? 'Yes' : 'No'}</dd><dt>Required Graph Scope</dt><dd>User.Read</dd><dt>Custom API Resource</dt><dd>{target || 'Not configured'}</dd></dl>{summary.missing.map(result => <p role="alert" key={result.id}>{result.message}</p>)}{summary.warnings.map(result => <p key={result.id}>{result.message}</p>)}{summary.unknown.map(result => <p key={result.id}>{result.message}</p>)}</>,
 *     diagnostics: <><p>Mount below the matching SPFx provider. Passive checks block auth popups/redirects while their public promise runs. Timeout does not cancel token acquisition; event guards detach when the overall check settles. Retry bypasses the token cache without changing passive mode. This diagnoses delegated scopes; it grants no access. Reset only clears this local completed view. Idle with no results is not a completed grant check.</p>{visible?.results && <Accordion collapsible><AccordionItem value="scopes"><AccordionHeader>Completed scope diagnostics</AccordionHeader><AccordionPanel><div className={styles.diagnostics}><ul>{visible.results.map(result => <li key={result.id}>{result.resourceName} / {result.scope}: {result.status}. {result.message}</li>)}</ul><pre>{JSON.stringify(visible.results, undefined, 2)}</pre></div></AccordionPanel></AccordionItem></Accordion>}</>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Manually check Graph User.Read and optional custom delegated scopes; compare a passive uncached retry and Reset while native work is pending.' };
 * ```
 */
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
  const requirementsKey = useMemo(
    () => createRequirementsKey(normalizedRequirements),
    [normalizedRequirements]
  );
  const stableRequirementsRef = useRef<readonly SPFxApiPermissionRequirement[]>(normalizedRequirements);
  const stableRequirementsKeyRef = useRef<string | undefined>(undefined);

  if (stableRequirementsKeyRef.current !== requirementsKey) {
    stableRequirementsKeyRef.current = requirementsKey;
    stableRequirementsRef.current = normalizedRequirements;
  }

  const stableRequirements = stableRequirementsRef.current;
  const service = useMemo(
    () => tokenProvider ? createSPFxApiPermissionPrecheckService(tokenProvider) : undefined,
    [tokenProvider]
  );
  const autoCheckKey = useMemo(
    () => createAutoCheckKey(requirementsKey, {
      autoCheck,
      mode,
      useCachedToken,
      validateAudience,
      timeoutMs
    }),
    [autoCheck, mode, requirementsKey, timeoutMs, useCachedToken, validateAudience]
  );

  const [isChecking, setIsChecking] = useState<boolean>(false);
  const [results, setResults] = useState<readonly SPFxApiPermissionCheckResult[]>([]);
  const isMountedRef = useRef<boolean>(true);
  const requestIdRef = useRef<number>(0);
  const autoCheckKeyRef = useRef<string | undefined>(undefined);
  const activeRequirementsKeyRef = useRef<string | undefined>(undefined);
  const activeTokenProviderRef = useRef(tokenProvider);

  useEffect(() => {
    isMountedRef.current = true;

    return () => {
      isMountedRef.current = false;
    };
  }, []);

  useEffect(() => {
    if (activeRequirementsKeyRef.current === undefined) {
      activeRequirementsKeyRef.current = requirementsKey;
      return;
    }

    if (activeRequirementsKeyRef.current === requirementsKey && activeTokenProviderRef.current === tokenProvider) {
      return;
    }

    activeRequirementsKeyRef.current = requirementsKey;
    activeTokenProviderRef.current = tokenProvider;
    requestIdRef.current += 1;
    autoCheckKeyRef.current = undefined;

    if (isMountedRef.current) {
      setIsChecking(false);
      setResults([]);
    }
  }, [requirementsKey, tokenProvider]);

  useEffect(() => {
    if (tokenProvider || isInitializing) {
      return;
    }

    const requestId = requestIdRef.current + 1;
    requestIdRef.current = requestId;
    const unavailableResults = createTokenProviderUnavailableResults(stableRequirements);

    if (isMountedRef.current && requestId === requestIdRef.current) {
      setIsChecking(false);
      setResults(unavailableResults);
    }
  }, [isInitializing, requirementsKey, stableRequirements, tokenProvider]);

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
            timeoutMs,
            sequentialResourceAcquisition: mode === 'passive'
          })
        )
        : createTokenProviderUnavailableResults(stableRequirements);

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
    service,
    stableRequirements,
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

function createRequirementsKey(
  requirements: readonly SPFxApiPermissionRequirement[]
): string {
  try {
    return JSON.stringify(requirements);
  } catch {
    return 'unserializable-requirements';
  }
}

function createAutoCheckKey(
  requirementsKey: string,
  options: Required<SPFxApiPermissionPrecheckOptions>
): string {
  try {
    return JSON.stringify({ requirementsKey, options });
  } catch {
    return 'unserializable-config';
  }
}
