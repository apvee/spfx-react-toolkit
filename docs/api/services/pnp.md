# PnP Generic Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxPnPService } from '@apvee/spfx-react-toolkit/services';
```


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

The caller owns the provided `SPFI` instance and its configuration.
