import { resolvePickersStrings } from './pickers-strings.internal';
import { PickerMessageError, projectPickerError } from './picker-message.internal';
import { useContext, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { useStableCallback } from '../../hooks/useStableCallback';
import type { PickerValue } from './picker-values.types';
import type { PickerOption, PickerProps, PickerSelectionEvent, PrincipalPickerInitialPrincipal, UserIdentityResolution } from './picker.types';
import type { PickerSource, PickerSourcePage, PickerSourceRequest, PickerHydrationInput, PickerPersonalSuggestions } from './picker-source.types.internal';
import { createPickerRequestCache } from './picker-request-cache.internal';
import { PickerRuntimeContext, PickerSelectedValueContext, PickerAuthoritativeEpochContext } from './picker-runtime-context.internal';
import { arePickerValuesCompatible, canonicalPickerIdentity, decodePickerAlias, pickerIdentityAliases, reconcilePickerAliases } from './picker-identity.internal';
import { validatePickerOption, copyPickerValue } from './picker-values-validation.internal';
import type { PickerAliasState } from './picker-identity.internal';
import { composePickerSuggestions } from './picker-suggestions.internal';

interface RuntimeDefaults { readonly multiselect: boolean; readonly initialCatalogLimit?: number; }
interface PickerSearchFeedback {
  readonly phase: 'idle' | 'pending' | 'complete' | 'partial' | 'error';
  readonly matchesVisibleChoices: boolean;
  readonly error?: Error;
  readonly warning?: string;
}
interface PrincipalDefaults { readonly defaultPrincipals?: readonly PrincipalPickerInitialPrincipal[]; readonly userIdentity?: UserIdentityResolution; }
const errorOf = (error: unknown): Error => error instanceof Error ? error : new Error(String(error));
const checkInteger = (value: number, name: string, min: number, max = Number.MAX_SAFE_INTEGER): number => {
  if (!Number.isSafeInteger(value) || value < min || value > max) throw new RangeError(`${name} must be an integer from ${min} to ${max}.`);
  return value;
};
/** Display metadata only; arbitrary fields never become request identity or serialized selection. */
function modelSignature(options: readonly PickerOption[]): string {
  return JSON.stringify(options.map(option => {
    try {
      const valid = validatePickerOption(option);
      return [valid.key, valid.text, copyPickerValue(valid.value), Object.entries(valid).filter(([key, value]) => !['key', 'text', 'value', 'fields'].includes(key) && ['string', 'number', 'boolean'].includes(typeof value))];
    } catch { return ['invalid']; }
  }));
}

/** Pure projections with effect-owned discovery, hydration and explicit transactional confirmation. */
export function usePickerRuntime<V extends PickerValue, T extends PickerOption<V>>(props: PickerProps<T> & PrincipalDefaults, source: PickerSource<V, T>, defaults: RuntimeDefaults) {
  const strings = resolvePickersStrings(props.strings);
  const projectedErrors = new Map<Error, Error>();
  const projectError = (error: Error | undefined): Error | undefined => {
    if (error === undefined) return undefined;
    const projected = projectedErrors.get(error) ?? projectPickerError(error, strings) ?? error;
    projectedErrors.set(error, projected);
    return projected;
  };
  const catalogLimit = defaults.initialCatalogLimit === undefined ? undefined : checkInteger(defaults.initialCatalogLimit, 'initialCatalogLimit', 1, 200);
  const minimum = checkInteger(props.minimumSearchLength ?? 2, 'minimumSearchLength', 2);
  const debounce = checkInteger(props.debounceMs ?? 300, 'debounceMs', 0);
  const pageSize = checkInteger(props.suggestionsLimit ?? 50, 'suggestionsLimit', 1, 200);
  const initialLimit = checkInteger(props.initialSuggestions?.limit ?? 5, 'initialSuggestions.limit', 1, 5);
  const multi = props.multiselect ?? defaults.multiselect;
  const limit = props.selectionLimit === undefined ? (multi ? Infinity : 1) : Math.min(checkInteger(props.selectionLimit, 'selectionLimit', 1), multi ? Infinity : 1);
  if (props.selectedOptions === undefined && props.defaultPrincipals !== undefined && props.defaultSelectedOptions !== undefined) throw new Error('defaultPrincipals and defaultSelectedOptions are mutually exclusive.');
  const sharedCache = useContext(PickerRuntimeContext);
  const knownReferences = useContext(PickerSelectedValueContext);
  const authoritativeEpoch = useContext(PickerAuthoritativeEpochContext);
  const ownCache = useMemo(() => createPickerRequestCache(source.context.owner), [source.context.owner]);
  const cache = sharedCache ?? ownCache;
  useEffect(() => () => ownCache.dispose(), [ownCache]);
  const context = source.context;
  const graphIdentity = context.graphClientIdentity ?? context.graphClient;
  const binding = useMemo(() => ({ source, partition: JSON.stringify([source.identity, context.tenantId, context.userId, context.webUrl, context.siteCollectionUrl,
    cache.identityKey(context.owner), cache.identityKey(context.spClient), cache.identityKey(graphIdentity)]) + ':' }),
  [source.identity, context.owner, context.tenantId, context.userId, context.webUrl, context.siteCollectionUrl, context.spClient, graphIdentity, cache]);
  const refreshPartition = JSON.stringify(['picker-source', source.family ?? source.identity, cache.identityKey(context.owner), context.tenantId, context.userId]) + ':';
  const cachedRequest = <R,>(key: string, load: () => Promise<R>): Promise<R> => cache.request(key, load, refreshPartition);
  const [retryRevision, setRetryRevision] = useState(0);
  const verificationEpoch = useMemo(() => ({}), [binding, authoritativeEpoch, retryRevision]);
  const [query, setQuery] = useState('');
  const searchText = source.caseInsensitive ? query.trim().toLowerCase() : query.trim();
  const [open, setOpen] = useState(false);
  const [keys, setKeys] = useState<readonly string[]>(() => props.defaultSelectedOptions ?? props.defaultPrincipals?.map((principal, index) => principal.kind === 'displayOnly' ? principal.key : `principal-initial:${index}`) ?? []);
  const initialPrincipals = useRef(new Map(props.selectedOptions === undefined ? keys.map((key, index) => [key, props.defaultPrincipals?.[index]] as const) : []));
  const selectedOptions = props.selectedOptions ?? keys;
  const selectedSignature = JSON.stringify(selectedOptions);
  const seeds = [...(props.initialOptions ?? []), ...(props.availableOptions ?? [])];
  const seedSignature = modelSignature(seeds);
  const referenceSignature = JSON.stringify(selectedOptions.map(key => {
    const value = knownReferences?.get(key); return value === undefined ? key : [key, canonicalPickerIdentity(value, context)];
  }));
  const [models, setModels] = useState<PickerAliasState<T>>(() => reconcilePickerAliases(undefined, seeds, context));
  const projectedModels = useMemo(() => reconcilePickerAliases(models, seeds, context), [models, seedSignature, context.tenantId]);
  const modeledSelection = selectedOptions.flatMap(key => { const option = projectedModels.options.get(key); return option === undefined ? [] : [option]; });
  const missingModels = selectedOptions.filter(key => !projectedModels.options.has(key));
  const verificationIdentity = (option: T): string => JSON.stringify(copyPickerValue(option.value), (key, value: unknown) => key === 'text' ? undefined : value);
  const [selectedScope, setSelectedScope] = useState<{ readonly epoch: object; readonly identities: ReadonlySet<string>; readonly settled: ReadonlySet<string>; readonly needed: ReadonlySet<string> }>();
  const selectedModelSignature = modelSignature(modeledSelection);
  const selectedItems = source.preflight ? modeledSelection.filter(option => selectedScope?.epoch === verificationEpoch && selectedScope.identities.has(verificationIdentity(option))) : modeledSelection;
  const unresolvedOptions = selectedOptions.filter(key => !selectedItems.some(option => option.key === key));
  const [hydrated, setHydrated] = useState<{ readonly epoch: object; readonly settled: ReadonlySet<string>; readonly errors: ReadonlyMap<string, Error> }>();
  const hydrationIdentity = (key: string): string => JSON.stringify([key, knownReferences?.get(key) ?? initialPrincipals.current.get(key) ?? decodePickerAlias(key)], (member, value: unknown) => member === 'text' ? undefined : value);
  const pendingHydration = missingModels.filter(key => hydrated?.epoch !== verificationEpoch || !hydrated.settled.has(hydrationIdentity(key)));
  const pendingScope = source.preflight ? modeledSelection.filter(option => selectedScope?.epoch !== verificationEpoch || !selectedScope.settled.has(verificationIdentity(option))).map(option => option.key) : [];
  const pendingKeys = new Set([...pendingHydration, ...pendingScope]);
  const terminalUnresolved = unresolvedOptions.filter(key => !pendingKeys.has(key));
  const [catalog, setCatalog] = useState<{ readonly binding: typeof binding; readonly revision: number; readonly options: readonly T[]; readonly partial: boolean; readonly error?: Error; readonly errors?: PickerSourcePage<T>['errors'] }>();
  const [remote, setRemote] = useState<{ readonly binding: typeof binding; readonly query: string; readonly pageSize: number; readonly revision: number; readonly firstPageRequest: Promise<PickerSourcePage<T>>; readonly page?: PickerSourcePage<T>; readonly loading?: boolean; readonly error?: Error }>();
  const [preflight, setPreflight] = useState<{ readonly binding: typeof binding; readonly revision: number; readonly permitted: ReadonlySet<string>; readonly settled: ReadonlySet<string> }>();
  const [personal, setPersonal] = useState<{ readonly binding: typeof binding; readonly revision: number; readonly data: PickerPersonalSuggestions<T>; readonly error?: Error }>();
  const [failure, setFailure] = useState<Error>();
  const [validation, setValidation] = useState<string | { readonly messageKey: keyof typeof strings.validation }>();
  const [unresolvedLabels, setUnresolvedLabels] = useState<ReadonlyMap<string, string>>(() => new Map(props.defaultPrincipals?.flatMap((principal, index) => principal.text === undefined ? [] : [[principal.kind === 'displayOnly' ? principal.key : `principal-initial:${index}`, principal.text] as const])));
  // Native-only initial presentation: a changed source/reference/epoch closes this cycle permanently.
  // Readiness/client initialization is deliberately absent: Graph-dependent hydration may still be waiting.
  const [initialHydration] = useState(() => ({ sourceIdentity: source.identity, referenceSignature, authoritativeEpoch, retryRevision }));
  const [initialHydrationClosed, setInitialHydrationClosed] = useState(knownReferences === undefined);
  const nativeSelectionLabel = (key: string): string | undefined =>
    [projectedModels.options.get(key)?.text, unresolvedLabels.get(key), knownReferences?.get(key)?.text].find(label => label !== undefined && label.trim().length > 0);
  const hasSelectionLabel = selectedOptions.some(key => nativeSelectionLabel(key) !== undefined);
  const initialHydrationPending = knownReferences !== undefined && !initialHydrationClosed && !hasSelectionLabel && pendingHydration.length > 0 &&
    initialHydration.sourceIdentity === source.identity && initialHydration.referenceSignature === referenceSignature &&
    initialHydration.authoritativeEpoch === authoritativeEpoch && initialHydration.retryRevision === retryRevision;
  useLayoutEffect(() => { if (!initialHydrationPending && !initialHydrationClosed) setInitialHydrationClosed(true); }, [initialHydrationPending, initialHydrationClosed]);
  const [confirming, setConfirming] = useState(false);
  const active = useRef({ binding, mounted: false, action: 0 });
  useLayoutEffect(() => {
    active.current = { binding, mounted: true, action: active.current.action + 1 };
    setConfirming(false);
    return () => { active.current.mounted = false; active.current.action++; };
  }, [binding, selectedSignature, multi, limit, props.validateSelection, authoritativeEpoch]);
  const committedQuery = useRef(query.trim());
  useLayoutEffect(() => { committedQuery.current = query.trim(); }, [query]);
  const discoveryGeneration = useRef(0);
  useLayoutEffect(() => { discoveryGeneration.current++; }, [binding, searchText, pageSize, retryRevision]);
  const validationGeneration = useRef(0);
  const request = (text = query, continuation?: string): PickerSourceRequest => {
    const generation = active.current.action;
    return { query: source.caseInsensitive ? text.trim().toLowerCase() : text.trim(), pageSize, continuation, generation, context, cache, partition: refreshPartition,
      isCurrent: () => active.current.mounted && active.current.binding === binding && active.current.action === generation };
  };
  const report = useStableCallback((error: Error, operation: 'load' | 'search' | 'resolve' | 'validate', suggestionSource?: 'people' | 'memberOf' | 'currentUser' | 'currentSite' | 'hubSite', transport?: 'sharePoint' | 'graph') => {
    const errorTransport = 'pickerSource' in error && (error.pickerSource === 'graph' || error.pickerSource === 'sharePoint') ? error.pickerSource : undefined;
    props.onError?.(projectError(error) ?? error, { source: transport ?? errorTransport ?? (suggestionSource === 'people' || suggestionSource === 'memberOf' ? 'graph' : source.transport), operation, suggestionSource });
  });
  const adopt = useStableCallback((options: readonly T[]) => {
    setModels(previous => {
      const next = reconcilePickerAliases(previous, [...options, ...seeds], context);
      const retained = new Map(next.options);
      // Preserve every selected model, bounded working set for unselected candidates.
      const protectedKeys = new Set(selectedOptions);
      for (const key of retained.keys()) if (retained.size > protectedKeys.size + 128 && !protectedKeys.has(key)) retained.delete(key);
      return { options: retained, errors: next.errors };
    });
  });
  useEffect(() => { adopt(seeds); }, [seedSignature, binding, adopt]);
  useEffect(() => {
    if (!source.ready || pendingHydration.length === 0) return;
    let disposed = false;
    const inputs: PickerHydrationInput<V>[] = pendingHydration.map(key => {
      const principal = initialPrincipals.current.get(key);
      if (principal !== undefined) return { kind: 'principalReference', key, reference: principal };
      const value = knownReferences?.get(key) ?? decodePickerAlias(key);
      // The adapter validates domain before resolving a decoded generic reference.
      return value === undefined ? { kind: 'key', key } : { kind: 'value', key, value: value as V };
    });
    const cacheKey = binding.partition + 'hydrate:' + JSON.stringify(inputs.map(input => input.kind === 'value' ? { ...input, value: { ...copyPickerValue(input.value), text: undefined } } : input));
    const settle = (failures: ReadonlyMap<string, Error>): void => setHydrated(previous => {
      const current = previous?.epoch === verificationEpoch ? previous : undefined;
      const errors = new Map(current?.errors);
      inputs.forEach(input => { const error = failures.get(input.key); if (error === undefined) errors.delete(input.key); else errors.set(input.key, error); });
      return { epoch: verificationEpoch, settled: new Set([...(current?.settled ?? []), ...inputs.map(input => hydrationIdentity(input.key))]), errors };
    });
    void cachedRequest(cacheKey, () => binding.source.hydrate(inputs, request(''))).then(result => {
      if (!disposed) {
        adopt(result.resolved.map(entry => ({ ...entry.option, key: entry.key })));
        const errors = new Map(result.unresolved.flatMap(entry => entry.error === undefined ? [] : [[entry.key, entry.error] as const]));
        settle(errors);
        new Set(errors.values()).forEach(error => report(error, 'resolve'));
        setUnresolvedLabels(previous => new Map([...previous, ...result.unresolved.flatMap(entry => entry.label === undefined ? [] : [[entry.key, entry.label] as const])]));
      }
    }, error => { if (!disposed) { const failure = errorOf(error); settle(new Map(inputs.map(input => [input.key, failure]))); report(failure, 'resolve'); } });
    return () => { disposed = true; };
  }, [binding, source.ready, selectedSignature, referenceSignature, seedSignature, retryRevision, verificationEpoch, adopt, report]);
  const searchCacheKey = binding.partition + 'search:' + JSON.stringify([searchText, pageSize]);
  const cachedSearchRequest = cache.peek(searchCacheKey);
  useEffect(() => {
    if (!source.ready || !open || props.searchable === false || searchText.length < minimum) return;
    const text = searchText;
    // This private key is populated only by this source's typed first-page request.
    const currentRequest = cache.peek(searchCacheKey, refreshPartition) as Promise<PickerSourcePage<T>> | undefined;
    // Retained pages belong to one exact request; a sibling can replace the cache entry while closed.
    if (remote?.binding === binding && remote.query === text && remote.pageSize === pageSize && remote.revision === retryRevision &&
      (remote.page !== undefined || remote.error !== undefined) && currentRequest === remote.firstPageRequest) return;
    // Invalidate obsolete continuations before the debounce interval, while retaining valid same-entry pages.
    if (currentRequest !== remote?.firstPageRequest) discoveryGeneration.current++;
    let disposed = false;
    const subscribe = (firstPageRequest: Promise<PickerSourcePage<T>>): void => {
      setRemote({ binding, query: text, pageSize, revision: retryRevision, firstPageRequest, loading: true });
      void firstPageRequest.then(page => {
        if (!disposed) { adopt(page.options); setRemote({ binding, query: text, pageSize, revision: retryRevision, firstPageRequest, page }); page.errors?.forEach(entry => report(entry.error, 'search', undefined, entry.source)); }
      }, error => { if (!disposed) { const failure = errorOf(error); setRemote({ binding, query: text, pageSize, revision: retryRevision, firstPageRequest, error: failure }); report(failure, 'search'); } });
    };
    const timer = currentRequest === undefined ? setTimeout(() => subscribe(cachedRequest(searchCacheKey, () => binding.source.discover(request(text)))), debounce) : undefined;
    if (currentRequest !== undefined) subscribe(currentRequest);
    return () => { disposed = true; if (timer !== undefined) clearTimeout(timer); };
  }, [binding, source.ready, open, minimum, debounce, pageSize, props.searchable, searchText, retryRevision, cachedSearchRequest, adopt, report]);
  const suggestionQuery = source.persistentSuggestions ? '' : query;
  useEffect(() => {
    if (!source.ready || !open || suggestionQuery.trim().length > 0 || props.showDefaultSuggestions === false || binding.source.personalSuggestions === undefined) return;
    let disposed = false;
    const load = binding.source.personalSuggestions;
    void cachedRequest(binding.partition + 'personal', () => load(request(''))).then(data => {
      if (!disposed) {
        adopt([...(data.currentSite ?? []), ...(data.hubSite ?? []), ...(data.currentUser ?? []), ...(data.people ?? []), ...(data.groups ?? [])]);
        setPersonal({ binding, revision: retryRevision, data });
        data.errors?.forEach(entry => report(entry.error, 'load', entry.source));
      }
    }, error => { if (!disposed) { const failure = errorOf(error); setPersonal({ binding, revision: retryRevision, data: { completeness: 'partial' }, error: failure }); report(failure, 'load'); } });
    return () => { disposed = true; };
  }, [binding, source.ready, open, suggestionQuery, props.showDefaultSuggestions, retryRevision, adopt, report]);
  const previousRefresh = useRef(props.refreshKey);
  useEffect(() => {
    if (previousRefresh.current === props.refreshKey) return;
    previousRefresh.current = props.refreshKey;
    cache.invalidate(refreshPartition);
    setRetryRevision(previous => previous + 1);
  }, [props.refreshKey, binding, cache, refreshPartition]);
  const searchActive = props.searchable !== false && searchText.length >= minimum;
  const visible = searchActive && remote?.binding === binding && remote.query === searchText && remote.pageSize === pageSize && remote.revision === retryRevision ? remote : undefined;
  // A new eligible query is pending immediately, including the debounce interval.
  const searchPending = searchActive && open && (visible === undefined || visible.loading === true || cachedSearchRequest !== visible.firstPageRequest);
  const visiblePersonal = (source.persistentSuggestions || query.trim() === '') && props.showDefaultSuggestions !== false && personal?.binding === binding && personal.revision === retryRevision ? personal : undefined;
  const personalData = visiblePersonal?.data;
  const personalPending = source.ready && open && suggestionQuery.trim() === '' && props.showDefaultSuggestions !== false &&
    binding.source.personalSuggestions !== undefined && visiblePersonal === undefined;
  const knownSelectedValues = selectedOptions.flatMap(key => { const value = projectedModels.options.get(key)?.value ?? knownReferences?.get(key) ?? decodePickerAlias(key); return value === undefined ? [] : [value]; });
  const selectedAliases = new Set(knownSelectedValues.flatMap(value => pickerIdentityAliases(value, context)));
  const acceptsOption = (option: T): boolean => {
    try {
      validatePickerOption(option);
      const retained = projectedModels.options.get(option.key);
      return source.accepts(option) && (retained === undefined || arePickerValuesCompatible(retained.value, option.value, context));
    } catch { return false; }
  };
  const permitted = (options: readonly T[], excludeSelected = true): readonly T[] => {
    const preferred = options.map(option => {
      const aliases = new Set(pickerIdentityAliases(option.value, context));
      return props.availableOptions?.find(consumer => acceptsOption(consumer) && pickerIdentityAliases(consumer.value, context).some(alias => aliases.has(alias))) ?? option;
    });
    const valid = preferred.filter(option => acceptsOption(option) && (!excludeSelected || !pickerIdentityAliases(option.value, context).some(alias => selectedAliases.has(alias))));
    const known = new Map(valid.map(option => [option.key, pickerIdentityAliases(option.value, context)]));
    const transformed = props.filterOptions?.(valid) ?? valid;
    const seen = new Set<string>();
    return transformed.filter(option => {
      let identity: readonly string[];
      try { validatePickerOption(option); identity = pickerIdentityAliases(option.value, context); } catch { return false; }
      if (!known.has(option.key) || !identity.every(alias => known.get(option.key)?.includes(alias)) || !acceptsOption(option) || identity.some(alias => excludeSelected && selectedAliases.has(alias) || seen.has(alias))) return false;
      identity.forEach(alias => seen.add(alias)); return true;
    });
  };
  useEffect(() => {
    if (!source.preflight || !source.ready || modeledSelection.length === 0) return;
    let disposed = false;
    const load = async (): Promise<void> => {
      for (let offset = 0; offset < modeledSelection.length && !disposed; offset += 20) {
        await Promise.all(modeledSelection.slice(offset, offset + 20).map(async option => {
          const identity = verificationIdentity(option);
          let accepted = false, needsIdentity = false;
          try {
            const verified = await cachedRequest(binding.partition + 'eligible:' + identity, () => binding.source.verifyCandidate(option, request('')));
            accepted = verified !== undefined;
            const needsResolution = binding.source.needsIdentityResolution;
            if (accepted && needsResolution !== undefined) needsIdentity = await cachedRequest(binding.partition + 'requiredIDs:' + identity, () => needsResolution(option, request('')));
          } catch (error) { needsIdentity = accepted; if (!disposed) report(errorOf(error), 'resolve'); }
          if (!disposed) setSelectedScope(previous => {
            const current = previous?.epoch === verificationEpoch ? previous : undefined;
            const identities = new Set(current?.identities), settled = new Set(current?.settled), needed = new Set(current?.needed);
            settled.add(identity); if (accepted) identities.add(identity); else identities.delete(identity);
            if (needsIdentity) needed.add(identity); else needed.delete(identity);
            return { epoch: verificationEpoch, identities, settled, needed };
          });
        }));
      }
    };
    void load();
    return () => { disposed = true; };
  }, [binding, source.ready, selectedSignature, selectedModelSignature, retryRevision, verificationEpoch, report]);
  const consumerSignature = modelSignature(props.availableOptions ?? []);
  const menuCandidates = (props.availableOptions ?? []).filter(option => acceptsOption(option) && !pickerIdentityAliases(option.value, context).some(alias => selectedAliases.has(alias)));
  const currentPreflight = preflight?.binding === binding && preflight.revision === retryRevision ? preflight : undefined;
  const preflightPending = source.preflight && source.ready && open && query.trim() === '' && props.showDefaultSuggestions !== false &&
    menuCandidates.some(option => !currentPreflight?.settled.has(verificationIdentity(option)));
  useEffect(() => {
    if (!source.preflight || !source.ready || !open || query.trim() !== '' || props.showDefaultSuggestions === false) return;
    let disposed = false;
    const candidates = menuCandidates;
    const load = async (): Promise<void> => {
      const permitted = new Set<string>();
      // Bound concurrency, not consumer set size. Scope filtering precedes the shared menu cap.
      for (let offset = 0; offset < candidates.length && !disposed; offset += 20) {
        const results = await Promise.all(candidates.slice(offset, offset + 20).map(async option => {
          const identity = verificationIdentity(option);
          try {
            const verified = await cachedRequest(binding.partition + 'eligible:' + identity, () => binding.source.verifyCandidate(option, request('')));
            return verified === undefined ? undefined : identity;
          } catch (error) { if (!disposed) report(errorOf(error), 'load'); return undefined; }
        }));
        results.forEach(identity => { if (identity !== undefined) permitted.add(identity); });
      }
      if (!disposed) setPreflight({ binding, revision: retryRevision, permitted, settled: new Set(candidates.map(verificationIdentity)) });
    };
    void load();
    return () => { disposed = true; };
  }, [binding, source.ready, open, query, consumerSignature, selectedSignature, props.showDefaultSuggestions, retryRevision, report]);
  const consumerOptions = source.preflight ? (props.availableOptions ?? []).filter(option => {
    try { return currentPreflight?.permitted.has(verificationIdentity(option)); } catch { return false; }
  }) : props.availableOptions;
  const visibleCatalog = catalog?.binding === binding && catalog.revision === retryRevision ? catalog : undefined;
  const catalogEnabled = (catalogLimit !== undefined || source.completeCatalog === true) && props.showDefaultSuggestions !== false && !props.disabled;
  useEffect(() => {
    if (!catalogEnabled || !source.ready || !open || query.trim() !== '') return;
    let disposed = false;
    const load = async (): Promise<void> => {
      const options: T[] = [], errors: NonNullable<PickerSourcePage<T>['errors']>[number][] = [], cursors = new Set<string>();
      let continuation: string | undefined, incomplete = false;
      try {
        for (let page = 0; page < 5; page++) {
          // Stable native page size reuses cached pages when selection or the display cap changes.
          const size = 200;
          const cursor = continuation;
          const key = binding.partition + 'catalog:' + JSON.stringify([size, cursor]);
          const result = await cachedRequest(key, () => binding.source.discover({ ...request('', cursor), pageSize: size }));
          options.push(...result.options); errors.push(...(result.errors ?? []));
          continuation = result.continuation;
          incomplete = result.completeness === 'partial';
          if (disposed) return;
          if (permitted([...(consumerOptions ?? []), ...options]).length > (catalogLimit ?? 30) || continuation === undefined) break;
          if (cursors.has(continuation)) throw new Error('Repeated catalog continuation.');
          cursors.add(continuation);
        }
        if (!disposed) { adopt(options); setCatalog({ binding, revision: retryRevision, options, partial: incomplete || continuation !== undefined, errors }); errors.forEach(entry => report(entry.error, 'load', undefined, entry.source)); }
      } catch (error) { if (!disposed) { setCatalog({ binding, revision: retryRevision, options, partial: true, error: errorOf(error) }); report(errorOf(error), 'load'); } }
    };
    void load();
    return () => { disposed = true; };
  }, [binding, source.ready, catalogEnabled, catalogLimit, open, query, selectedSignature, consumerSignature, props.filterOptions, retryRevision, adopt, report]);
  // Contextual sites are fixed navigation choices, outside the ordinary five suggestions.
  // Reuse the canonical guards and consumer model precedence before composing their sections.
  const pinnedSections = source.persistentSuggestions ? composePickerSuggestions({ strings,
    currentSite: permitted(personalData?.currentSite ?? []), hubSite: permitted(personalData?.hubSite ?? []),
    selected: selectedItems, selectedValues: knownSelectedValues, context, accepts: acceptsOption, showDefaultSuggestions: props.showDefaultSuggestions,
  }) : [];
  const pinnedValues = pinnedSections.flatMap(section => section.options.map(option => option.value));
  const pinnedAliases = new Set(pinnedValues.flatMap(value => pickerIdentityAliases(value, context)));
  const sections = composePickerSuggestions({ strings, consumer: consumerOptions, catalog: catalogEnabled ? visibleCatalog?.options : undefined, catalogLimit, completeCatalog: source.completeCatalog,
    currentSite: source.persistentSuggestions ? undefined : personalData?.currentSite, hubSite: source.persistentSuggestions ? undefined : personalData?.hubSite,
    currentUser: personalData?.currentUser, people: personalData?.people, groups: personalData?.groups,
    selected: selectedItems, selectedValues: [...knownSelectedValues, ...pinnedValues], context, accepts: acceptsOption, filterOptions: props.filterOptions, limit: initialLimit, showDefaultSuggestions: props.showDefaultSuggestions });
  const catalogMore = catalogEnabled && !source.completeCatalog && permitted([...(consumerOptions ?? []), ...(visibleCatalog?.options ?? [])]).length > (catalogLimit ?? 30);
  const catalogHint = query.trim() === '' && catalogMore ? strings.status.moreItemsAvailable : undefined;
  // Match/filter once before removing choices already represented by selections or fixed sections.
  const searchMatches = permitted(visible?.page?.options ?? [], false);
  const searchOptions = searchMatches.filter(option => !pickerIdentityAliases(option.value, context).some(alias => selectedAliases.has(alias) || pinnedAliases.has(alias)));
  const matchesVisibleChoices = searchMatches.some(option => pickerIdentityAliases(option.value, context).some(alias => selectedAliases.has(alias) || pinnedAliases.has(alias)));
  const emit = useStableCallback((event: PickerSelectionEvent, next: readonly string[], nextItems: readonly T[], reason: 'selection' | 'remove' | 'clear' | 'resolve', key?: string) => {
    if (props.selectedOptions === undefined) setKeys(next);
    props.onOptionSelect?.(event, { reason, selectedOptions: next, selectedItems: nextItems, selectedValues: nextItems.map(option => option.value),
      unresolvedOptions: next.filter(candidate => !nextItems.some(option => option.key === candidate)), optionValue: key, optionText: nextItems.find(option => option.key === key)?.text });
  });
  const reduceSelection = useStableCallback(async (event: PickerSelectionEvent, next: readonly string[], reason: 'remove' | 'clear', key?: string) => {
    if (props.disabled) return;
    const token = ++active.current.action, current = binding;
    validationGeneration.current++; setConfirming(false);
    const nextItems = selectedItems.filter(item => next.includes(item.key));
    try {
      // Without change validation the native gesture commits synchronously.
      const message = props.validationTrigger === 'blur' || props.validateSelection === undefined ? undefined : await props.validateSelection(nextItems);
      if (!active.current.mounted || active.current.binding !== current || active.current.action !== token) return;
      setValidation(message);
      if (message) return;
      if (reason === 'clear') setQuery('');
      emit(event, next, nextItems, reason, key);
    } catch (error) { if (active.current.mounted && active.current.binding === current && active.current.action === token) { setFailure(errorOf(error)); report(errorOf(error), 'validate'); } }
  });
  const remove = useStableCallback((event: PickerSelectionEvent, key: string) => reduceSelection(event, selectedOptions.filter(item => item !== key), 'remove', key));
  const clear = useStableCallback((event: PickerSelectionEvent) => reduceSelection(event, [], 'clear'));
  const select = useStableCallback(async (event: PickerSelectionEvent, key: string): Promise<boolean> => {
    if (props.disabled || confirming) return false;
    if (selectedOptions.includes(key)) { await remove(event, key); return false; }
    const candidate = projectedModels.options.get(key) ?? searchOptions.find(option => option.key === key) ?? [...pinnedSections, ...sections].flatMap(section => section.options).find(option => option.key === key);
    if (candidate === undefined || !source.accepts(candidate)) return false;
    const next = multi ? [...selectedOptions, key] : [key];
    if (next.length > limit) { setValidation({ messageKey: 'selectionLimitReached' }); return false; }
    const token = ++active.current.action;
    validationGeneration.current++;
    const current = active.current.binding;
    setConfirming(true); setFailure(undefined);
    try {
      const verified = await source.verifyCandidate(candidate, request());
      if (!active.current.mounted || active.current.binding !== current || active.current.action !== token) return false;
      if (verified === undefined) throw new PickerMessageError({ group: 'errors', key: 'candidateOutOfScope' });
      const confirmed = source.confirm === undefined ? [verified] : await source.confirm([verified], request(), 'selection');
      if (!active.current.mounted || active.current.binding !== current || active.current.action !== token) return false;
      const replacement = confirmed[0];
      if (replacement === undefined) throw new PickerMessageError({ group: 'errors', key: 'incompleteCandidateResolution' });
      const nextItems = next.flatMap(item => item === key ? [{ ...replacement, key }] : selectedItems.filter(option => option.key === item));
      const message = props.validationTrigger === 'blur' ? undefined : await props.validateSelection?.(nextItems);
      if (!active.current.mounted || active.current.binding !== current || active.current.action !== token) return false;
      setValidation(message);
      if (message) return false;
      adopt(nextItems); setQuery('');
      emit(event, next, nextItems, 'selection', key);
      return true;
    } catch (error) {
      if (active.current.mounted && active.current.binding === current && active.current.action === token) { const failure = errorOf(error); setFailure(failure); report(failure, 'validate'); }
      return false;
    } finally { if (active.current.mounted && active.current.binding === current && active.current.action === token) setConfirming(false); }
  });
  const resolve = useStableCallback(async (event: PickerSelectionEvent) => {
    if (props.disabled || confirming || source.confirm === undefined || unresolvedOptions.length > 0) return;
    const token = ++active.current.action, current = active.current.binding;
    setConfirming(true);
    try {
      const next = await source.confirm(selectedItems, request(), 'explicitResolution');
      const message = await props.validateSelection?.(next);
      if (!active.current.mounted || active.current.binding !== current || active.current.action !== token) return;
      setValidation(message);
      if (!message) { adopt(next); emit(event, selectedOptions, next, 'resolve'); }
    } catch (error) { if (active.current.mounted && active.current.binding === current && active.current.action === token) { setFailure(errorOf(error)); report(errorOf(error), 'resolve'); } }
    finally { if (active.current.mounted && active.current.binding === current && active.current.action === token) setConfirming(false); }
  });
  const loadMore = useStableCallback(async () => {
    const previousPage = visible?.page;
    if (visible === undefined || previousPage?.continuation === undefined || visible.loading) return;
    const continuation = previousPage.continuation;
    const generation = discoveryGeneration.current;
    const before = visible, text = searchText, current = binding;
    setRemote({ ...before, loading: true });
    try {
      const page = await cachedRequest(binding.partition + 'search:' + JSON.stringify([text, pageSize, continuation]), () => source.discover(request(text, continuation)));
      if (!active.current.mounted || active.current.binding !== current || (source.caseInsensitive ? committedQuery.current.toLowerCase() : committedQuery.current) !== text || discoveryGeneration.current !== generation) return;
      adopt(page.options); setRemote({ binding, query: text, pageSize, revision: retryRevision, firstPageRequest: before.firstPageRequest, page: { ...page, options: [...previousPage.options, ...page.options], errors: [...(previousPage.errors ?? []), ...(page.errors ?? [])] } });
      page.errors?.forEach(entry => report(entry.error, 'search', undefined, entry.source));
    } catch (error) { if (active.current.mounted && active.current.binding === current && (source.caseInsensitive ? committedQuery.current.toLowerCase() : committedQuery.current) === text && discoveryGeneration.current === generation) { setRemote({ ...before, error: errorOf(error) }); report(errorOf(error), 'search'); } }
  });
  const retry = useStableCallback(() => { cache.invalidate(refreshPartition); setFailure(undefined); setRetryRevision(previous => previous + 1); });
  const blur = useStableCallback(async () => {
    if (props.validationTrigger !== 'blur' || confirming) return;
    const token = ++validationGeneration.current, action = active.current.action, current = binding;
    try {
      const message = await props.validateSelection?.(selectedItems);
      if (active.current.mounted && active.current.binding === current && active.current.action === action && validationGeneration.current === token) setValidation(message);
    } catch (error) {
      if (active.current.mounted && active.current.binding === current && active.current.action === action && validationGeneration.current === token) { setFailure(errorOf(error)); report(errorOf(error), 'validate'); }
    }
  });
  const missingIdentity = (selectedScope?.epoch === verificationEpoch && selectedItems.some(option => selectedScope.needed.has(verificationIdentity(option)))) || selectedItems.some(option => option.value.kind === 'user' && ((props.userIdentity?.mode === 'entra' || props.userIdentity?.mode === 'both') && option.value.entra === undefined ||
    (props.userIdentity?.mode === 'sharePoint' || props.userIdentity?.mode === 'both') && option.value.sharePoint === undefined));
  const activeCatalog = query.trim() === '' && catalogEnabled ? visibleCatalog : undefined;
  const sourceErrors = visible?.page?.errors ?? activeCatalog?.errors ?? [];
  const sourceFailure = sourceErrors.find(entry => entry.retryable !== false)?.error ?? activeCatalog?.error ?? visiblePersonal?.error ?? personalData?.errors?.[0]?.error;
  const warnings = [...new Set(sourceErrors.filter(entry => entry.retryable === false).map(entry => projectError(entry.error)?.message))];
  const warning = warnings.length > 0 ? warnings.join(' ') : sourceFailure === undefined && (visible?.page?.completeness === 'partial' && visible.page.continuation === undefined || personalData?.completeness === 'partial' || activeCatalog?.partial && !catalogMore) ? strings.status.partialSearch : undefined;
  const contextualError = visiblePersonal?.error ?? personalData?.errors?.[0]?.error;
  const contextualWarning = contextualError === undefined && personalData?.completeness === 'partial' ? strings.status.partialSuggestions : undefined;
  const searchError = visible?.error ?? visible?.page?.errors?.find(entry => entry.retryable !== false)?.error;
  const searchWarnings = [...new Set((visible?.page?.errors ?? []).filter(entry => entry.retryable === false).map(entry => projectError(entry.error)?.message))];
  const searchFeedback: PickerSearchFeedback = {
    phase: !searchActive ? 'idle' : searchPending ? 'pending' : searchError !== undefined ? 'error' : visible?.page?.completeness === 'partial' ? 'partial' : visible?.page === undefined ? 'idle' : 'complete',
    matchesVisibleChoices, error: projectError(searchError), warning: searchWarnings.length > 0 ? searchWarnings.join(' ') : undefined,
  };
  const hydrationFailure = hydrated?.epoch === verificationEpoch ? missingModels.map(key => hydrated.errors.get(key)).find(error => error !== undefined) : undefined;
  const selectionText = (key: string): string => knownReferences === undefined ?
    projectedModels.options.get(key)?.text ?? unresolvedLabels.get(key) ?? strings.status.unresolvedValue :
    nativeSelectionLabel(key) ?? (pendingHydration.includes(key) ? strings.status.loading : strings.status.unresolvedValue);
  return { query, setQuery, open, setOpen, initialHydrationPending, selectedOptions, selectedItems, unresolvedOptions, pinnedSections, sections, searchOptions, searchFeedback, contextualError: projectError(contextualError), contextualWarning, confirming, selectionText, catalogHint, canResolve: source.confirm !== undefined && missingIdentity,
    loading: confirming || pendingKeys.size > 0 || searchPending || preflightPending || personalPending || (catalogEnabled && open && query.trim() === '' && source.ready && visibleCatalog === undefined), warning, error: projectError(failure ?? visible?.error ?? projectedModels.errors[0] ?? hydrationFailure ?? sourceFailure),
    partial: visible?.page?.completeness === 'partial' || personalData?.completeness === 'partial', continuation: visible?.page?.continuation,
    validation: selectedOptions.length > limit ? strings.validation.selectionLimitExceeded : knownReferences !== undefined && hydrationFailure !== undefined ? projectError(hydrationFailure)?.message : terminalUnresolved.length > 0 ? strings.validation.unresolvedValues : selectedItems.some(option => !acceptsOption(option)) ? strings.validation.outOfScopeValues : missingIdentity ? strings.validation.resolveRequiredUserIds : typeof validation === 'string' ? validation : validation === undefined ? undefined : strings.validation[validation.messageKey],
    minimum, select, remove, clear, resolve, retry, loadMore, blur };
}
