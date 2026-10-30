import { PickerMessageError } from './picker-message.internal';
import type { PickerOption } from './picker.types';
import type { PickerValue } from './picker-values.types';
import type { PickerSource, PickerSourceContext, PickerSourcePage, PickerSourceRequest } from './picker-source.types.internal';
import { canonicalPickerIdentity, decodePickerAlias } from './picker-identity.internal';
import { copyPickerValue } from './picker-values-validation.internal';

/** Shared read-only resolution/hydration semantics of the metadata domains, including unresolved aliases. */
export function createMetadataPickerSource<V extends PickerValue, T extends PickerOption<V>>(config: {
  readonly identity: string;
  readonly context: PickerSourceContext;
  readonly completeCatalog?: boolean;
  /** Optional ordering applied once after every native page has been collected. */
  readonly compare?: (first: T, second: T) => number;
  readonly kinds: readonly V['kind'][];
  readonly accepts: (option: T) => boolean;
  readonly discover: (request: PickerSourceRequest) => Promise<PickerSourcePage<T>>;
  readonly read: (value: V, request: PickerSourceRequest, verify: boolean) => Promise<T | undefined>;
}): PickerSource<V, T> {
  const catalogKey = (request: PickerSourceRequest): string => {
    const context = request.context;
    return JSON.stringify(['complete-picker-catalog', config.identity, context.tenantId, context.userId, context.webUrl, context.siteCollectionUrl,
      request.cache.identityKey(context.owner), request.cache.identityKey(context.spClient ?? context.owner)]);
  };
  const source: PickerSource<V, T> = {
    preflight: true, caseInsensitive: true, family: config.kinds.join(':'), identity: config.identity, context: config.context, ready: config.context.spClient !== undefined, transport: 'sharePoint',
    completeCatalog: config.completeCatalog,
    accepts: config.accepts,
    async discover(request) {
      if (!config.completeCatalog) return config.discover({ ...request, query: request.query.toLowerCase() });
      const catalog = await request.cache.request(catalogKey(request), async () => {
        const options: T[] = [], errors: NonNullable<PickerSourcePage<T>['errors']>[number][] = [], seen = new Set<string>();
        let page: PickerSourcePage<T>;
        let continuation: string | undefined;
        do {
          const cursor = continuation;
          page = await config.discover({ ...request, query: '', pageSize: 200, continuation: cursor });
          options.push(...page.options); errors.push(...(page.errors ?? []));
          continuation = page.continuation;
          if (continuation !== undefined) {
            if (seen.has(continuation)) throw new Error('Repeated complete catalog continuation.');
            seen.add(continuation);
          }
        } while (continuation !== undefined);
        if (config.compare !== undefined) options.sort(config.compare);
        return { options, completeness: page.completeness, errors };
      }, request.partition);
      const query = request.query.trim().toLowerCase();
      return { ...catalog, options: query === '' ? catalog.options : catalog.options.filter(option => option.text.toLowerCase().startsWith(query)) };
    },
    async resolve(values, request) {
      const resolved: { value: V; option: T }[] = [], unresolved: { value: V; error?: Error }[] = [];
      for (const input of values) {
        try {
          const value = copyPickerValue(input);
          if (!config.kinds.includes(value.kind as V['kind'])) throw new Error('Reference domain does not match picker.');
          const option = await config.read(value as V, request, false);
          if (option === undefined) unresolved.push({ value: input });
          else resolved.push({ value: option.value, option });
        } catch (error) { unresolved.push({ value: input, error: error instanceof Error ? error : new PickerMessageError({ group: 'errors', key: 'resolutionFailed' }) }); }
      }
      return { resolved, unresolved };
    },
    async hydrate(inputs, request) {
      const resolved: { key: string; value: V; option: T }[] = [], unresolved: { key: string; label?: string; error?: Error }[] = [];
      for (const input of inputs) {
        const value = input.kind === 'value' ? input.value : input.kind === 'key' ? decodePickerAlias(input.key) : undefined;
        if (value === undefined || !config.kinds.includes(value.kind as V['kind'])) { unresolved.push({ key: input.key }); continue; }
        const result = await source.resolve([value as V], request);
        const entry = result.resolved[0];
        if (entry === undefined) unresolved.push({ key: input.key, label: value.text, error: result.unresolved[0]?.error });
        else resolved.push({ key: input.key, value: entry.value, option: { ...entry.option, key: input.key } });
      }
      return { resolved, unresolved };
    },
    async verifyCandidate(option, request) {
      if (!config.accepts(option)) return undefined;
      const identity = canonicalPickerIdentity(option.value, config.context);
      let verified: T | undefined;
      if (config.completeCatalog) {
        const key = catalogKey(request);
        // This private key is populated only by the complete catalog loader above.
        const pending = request.cache.peek(key, request.partition) as Promise<PickerSourcePage<T>> | undefined;
        if (pending !== undefined) {
          try {
            const catalog = await pending;
            if (request.cache.peek(key) === pending && catalog.completeness === 'complete' && !catalog.errors?.length) {
              verified = catalog.options.find(candidate => config.accepts(candidate) && canonicalPickerIdentity(candidate.value, config.context) === identity);
            }
          } catch { /* A catalog failure does not prevent an independent filtered entity read. */ }
        }
      }
      verified ??= await config.read(option.value, request, true);
      if (verified === undefined || canonicalPickerIdentity(verified.value, config.context) !== identity) return undefined;
      return { ...verified, key: option.key, text: option.text };
    },
  };
  return source;
}
