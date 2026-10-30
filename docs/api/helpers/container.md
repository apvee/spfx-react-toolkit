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

The required `width` is a CSS-pixel measurement supplied by the caller; this helper does not measure an element or need a provider. Thresholds are exclusive upper bounds, so `480` is medium and `640` is large. It performs no validation: negative values classify as small, and `NaN`/positive infinity fall through to xxxLarge. Supply finite nonnegative measurements for meaningful layout decisions.

```ts
const category = getSPFxContainerSize(720); // large
const boundary = getSPFxContainerSize(480); // medium
```

For live element measurement, use the [container hook](../hooks/theming.md#usespfxcontainersize). CSS container/viewport style scopes use their own inclusive thresholds; see [styles](styles.md#types-and-scopes).
