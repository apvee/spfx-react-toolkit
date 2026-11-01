## StageProps

Stage presentation constraints.

### Imports (9)

```ts
import type { StageProps } from '@fixture/render';
import type { StageProps } from '@fixture/render/controls';
```

### Type (5)

```ts
type StageProps = StageBase & (
  { presentation: 'dialog'; trigger?: never } |
  { presentation: 'popover'; trigger: string }
)
```

### Members (8)

#### StageProps branch 1

```ts
StageBase & { presentation: "dialog"; trigger?: undefined; }
```

| Member | Type | Requirement | Default | Description |
| --- | --- | --- | --- | --- |
| `StageProps.open` | `boolean \| undefined` | Optional | See [StageBase.open](stage.md#stagebase) | Whether the surface is open. Omitted means closed. Inherited from [StageBase.open](stage.md#stagebase). |
| `StageProps.presentation` | `"dialog"` | Required | — |  |
| `StageProps.trigger` | `never` | Forbidden | — |  |

#### StageProps branch 2

```ts
StageBase & { presentation: "popover"; trigger: string; }
```

| Member | Type | Requirement | Default | Description |
| --- | --- | --- | --- | --- |
| `StageProps.open` | `boolean \| undefined` | Optional | See [StageBase.open](stage.md#stagebase) | Whether the surface is open. Omitted means closed. Inherited from [StageBase.open](stage.md#stagebase). |
| `StageProps.presentation` | `"popover"` | Required | — |  |
| `StageProps.trigger` | `string` | Required | — |  |

### Related types (6)

- [StageBase](stage.md#stagebase)

### Source (10)

[src/controls.ts:13](https://github.com/fixture/render/blob/stable/src/controls.ts#L13)
