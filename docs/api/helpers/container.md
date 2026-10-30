# Container Helpers

[Helpers overview](../INDEX.md#helpers-and-styles) · [Non-React composition](../../guides/helpers-and-services.md)


Import:

```ts
import { getSPFxContainerSize } from '@apvee/spfx-react-toolkit';
```

## `getSPFxContainerSize`

```ts
function getSPFxContainerSize(
  width: number
): 'small' | 'medium' | 'large' | 'xLarge' | 'xxLarge' | 'xxxLarge';
```

| Width | Result |
|-------|--------|
| `< 480` | `small` |
| `480-639` | `medium` |
| `640-1023` | `large` |
| `1024-1365` | `xLarge` |
| `1366-1919` | `xxLarge` |
| `>= 1920` | `xxxLarge` |
