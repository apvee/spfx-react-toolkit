# PnP Generic Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

## `createSPFxPnPService`

```ts
function createSPFxPnPService(sp: SPFI): SPFxPnPService;
```

Required public type:

- `SPFxPnPService`

Returned methods:

| Method | Description |
|--------|-------------|
| `invoke(fn)` | Runs `fn(sp)` and returns its result |
| `batch(fn)` | Creates a batched `SPFI`, executes the batch, and returns the result of `fn` |

The `sp` property is the exact supplied client; the caller owns its configuration. The factory makes no request. `invoke` forwards callback failures. `batch` calls the callback to queue requests, executes the batch, then awaits its result. Queue work before awaiting its request results: requests queued after an `await` may miss execution. Batch execution failures reject, and successful writes are not rolled back.

## Usage

```ts
import '@pnp/sp/webs';
import { createSPFxPnPService } from '@apvee/spfx-react-toolkit/services';

const service = createSPFxPnPService(sp); // Authenticated SPFI supplied by the caller.
const currentWeb = await service.invoke(client => client.web());
const batchedWeb = await service.batch(client => client.web());
console.log(service.sp === sp, currentWeb, batchedWeb); // true
```

`SPFxPnPService.sp` is readonly and refers to the original client. Both callbacks have shape `<T>(sp: SPFI) => Promise<T>` and return `Promise<T>` through their service method. This module registers batching only. It does not acquire credentials, add caching, manage React state or register arbitrary caller features. A batch is not a transaction, and errors may follow successful server writes.
