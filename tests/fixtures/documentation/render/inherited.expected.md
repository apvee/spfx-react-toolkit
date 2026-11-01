## Derived

Inherits the open state.

### Imports

```ts
import type { Derived } from '@fixture/render';
import type { Derived } from '@fixture/render/controls';
```

### Type

```ts
interface Derived extends StageBase
```

### Members

| Member | Type | Requirement | Default | Description |
| --- | --- | --- | --- | --- |
| `Derived.label` | `string` | Required | — | Label content. |
| `Derived.open` | `boolean \| undefined` | Optional | See [StageBase.open](stage.md#stagebase) | Whether the surface is open. Omitted means closed. Inherited from [StageBase.open](stage.md#stagebase). |

### Related types

- [StageBase](stage.md#stagebase)

### Source

[src/controls.ts:40](https://github.com/fixture/render/blob/stable/src/controls.ts#L40)
