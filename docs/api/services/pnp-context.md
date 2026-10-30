# PnP Context Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxPnPContextService } from '@apvee/spfx-react-toolkit/services';
```


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
| `resolveSiteUrl(siteUrl?)` | Resolves undefined to the current web URL, server-relative URLs to the current origin, and trims trailing slash |
| `createSPFI(siteUrl?, config?)` | Creates a configured PnPjs `SPFI` instance |
| `getConfigKey(config?)` | Returns a JSON string key for memoizing config |

`createSPFI` throws when `spfxContext` is not available. Cache expiry uses `Date.now() + timeout` through PnPjs `expireFunc`. A custom `cache.keyFactory` is used directly; `getConfigKey()` is a JSON key and cannot represent function identity, so callers memoizing their own service configuration must track the key factory reference separately.
