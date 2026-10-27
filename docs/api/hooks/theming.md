# Theme and Container Hooks

Public hook signatures and result shapes match the library source. External types (React, SPFx and PnPjs) come from their respective packages. All hooks require a matching SPFx provider.

## useSPFxThemeInfo

```typescript
export function useSPFxThemeInfo(): IReadonlyTheme | undefined;

```

Returns the current SPFx ThemeProvider theme when available. The provider owns its theme subscription and removes it on unmount.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxThemeInfo.ts)

## useSPFxFluent9ThemeInfo

```typescript
export function useSPFxFluent9ThemeInfo(): SPFxFluent9ThemeInfo;
export interface SPFxFluent9ThemeInfo {
    readonly theme: Theme;
    readonly isTeams: boolean;
    readonly teamsTheme?: string;
}
```

Chooses a normalized application SDK theme when context acquisition succeeded and a theme is available, otherwise converts the SPFx theme, otherwise uses the light-theme fallback. SharePoint updates come from the existing provider subscription; Teams/Office/Outlook updates come from the live SDK dispatcher described in [environment hooks](./environment.md#usespfxteams).

The legacy `isTeams` field equals SDK context support, including Office/Outlook; it does not classify the actual host or guarantee that an SDK theme is available. Optional `teamsTheme` is the normalized default/dark/highContrast theme. Use `useSPFxEnvironmentInfo` for host classification. There is no `isDark` field. The SDK's single-handler ownership and lack of unregister also apply here.

[WebPartShell](../controls/web-part-shell.md) composes this hook with explicit theme > nearest ancestor Fluent tokens > host tokens, and explicit direction > actual ancestor Fluent direction > SPFx culture. It supplies the owner document but inherits the renderer; hosts must supply a matching Griffel renderer.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxFluent9ThemeInfo.ts)

## useSPFxContainerSize

```typescript
export function useSPFxContainerSize(): SPFxContainerSizeInfo;
export type SPFxContainerSize = 'small' | 'medium' | 'large' | 'xLarge' | 'xxLarge' | 'xxxLarge';
export interface SPFxContainerSizeInfo {
    readonly size: SPFxContainerSize;
    readonly isSmall: boolean;
    readonly isMedium: boolean;
    readonly isLarge: boolean;
    readonly isXLarge: boolean;
    readonly isXXLarge: boolean;
    readonly isXXXLarge: boolean;
    readonly width: number;
    readonly height: number;
}
```

Reports the provider's observed dimensions and a breakpoint: small below 480px, medium 480–639px, large 640–1023px, xLarge 1024–1365px, xxLarge 1366–1919px and xxxLarge from 1920px. The boolean flags correspond to the current breakpoint, not a cumulative minimum width.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxContainerSize.ts)

## useSPFxContainerInfo

```typescript
export function useSPFxContainerInfo(): SPFxContainerInfo;
export interface SPFxContainerInfo {
    readonly element: HTMLElement | undefined;
    readonly size: ContainerSize | undefined;
}
```

Returns the host container element and optional observed size. Some extension hosts have no container; both values can be undefined. Read width/height from `size`, not directly from this result.

```tsx
import * as React from 'react';
import { useSPFxContainerInfo } from '@apvee/spfx-react-toolkit';

function Dimensions() {
  const { size } = useSPFxContainerInfo();
  return <p>{size ? size.width + ' × ' + size.height : 'No container'}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxContainerInfo.ts)

## See Also

- [API index](../../INDEX.md)
- [Services](../services/INDEX.md)
- [SharePoint validation](../../SHAREPOINT-VALIDATION.md)
