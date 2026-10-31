# PnP Context Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

## `createSPFxPnPContextService`

```ts
function createSPFxPnPContextService(
  spfxContext: ISPFXContext | undefined,
  pageContext: SPFxPnPPageContextLike
): SPFxPnPContextService;
```

Required public types:

- `SPFxPnPContextServiceConfig`
- `SPFxPnPPageContextLike`
- `SPFxPnPContextService`

`SPFxPnPPageContextLike` requires:

```ts
interface SPFxPnPPageContextLike {
  web: { absoluteUrl: string };
}
```

`SPFxPnPContextServiceConfig` supports:

| Option | Purpose |
|--------|---------|
| `cache.enabled` | Enables PnPjs caching |
| `cache.storage` | `session` or `local`; defaults to `session` |
| `cache.timeout` | Cache timeout in milliseconds, applied as PnPjs expiry; defaults to `300000`, and `0` retains that legacy fallback |
| `cache.keyFactory` | Custom cache key builder |
| `batch.enabled` | Reserved config marker for callers that coordinate batching |
| `batch.maxRequests` | Reserved max request marker |
| `headers` | Extra headers injected into PnPjs requests |

Returned `SPFxPnPContextService` methods:

| Method | Description |
|--------|-------------|
| `resolveSiteUrl(siteUrl?)` | Resolves omitted/empty input to the current web unchanged; supplied server-relative URLs use the current origin, and supplied URLs lose one trailing slash |
| `createSPFI(siteUrl?, config?)` | Creates a configured PnPjs `SPFI` instance |
| `getConfigKey(config?)` | Returns a JSON string key for memoizing config |

`createSPFI` throws when `spfxContext` is not available. Cache expiry uses `Date.now() + timeout` through PnPjs `expireFunc`. A custom `cache.keyFactory` is used directly; `getConfigKey()` is a JSON key and cannot represent function identity, so callers memoizing their own service configuration must track the key factory reference separately.

## Usage and dependency ownership

```ts
const factory = createSPFxPnPContextService(context, context.pageContext);
const sp = factory.createSPFI('/sites/projects', {
  cache: { enabled: true, storage: 'session', timeout: 60000 },
  headers: { 'X-Application': 'reports' },
});
const identity = factory.getConfigKey({ cache: { enabled: false } });
```

Construction performs no requests. The caller owns SPFx authentication context and any `SPFI` clients returned. No React provider is required when calling this service directly, even though its missing-context error refers to the provider for hook consumers. This module registers PnP webs and batching; callbacks and direct PnP operations need explicit imports for their own feature modules. Configuration adds no cache unless `cache.enabled` is true, no headers unless supplied, and no batching behavior from the reserved `batch` metadata.

Omitted or empty `siteUrl` returns the current web string unchanged, including its trailing slash. A supplied URL loses one final slash only; no whitespace trimming, HTTP validation or site discovery occurs. Server-relative resolution parses the current web with `URL`, which can throw. `getConfigKey(undefined)` returns `'{}'`; it can throw for non-JSON-serializable configuration. The default cache key hashes the request URL and is not a collision-proof application identity.

## Configuration shape

```ts
interface SPFxPnPContextServiceConfig {
  cache?: {
    enabled: boolean;
    storage?: 'session' | 'local';
    timeout?: number;
    keyFactory?: (url: string) => string;
  };
  batch?: { enabled: boolean; maxRequests?: number };
  headers?: Record<string, string>;
}
```

`enabled` is required when its enclosing cache/batch object is supplied. The cache factory returns its key string to PnPjs unchanged; there is no toolkit validation or URL namespacing beyond the default URL hash. The reserved batch fields remain metadata only.
