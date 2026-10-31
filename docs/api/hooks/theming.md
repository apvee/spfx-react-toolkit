# Theme and Container Hooks

[React and SPFx hooks](../INDEX.md#react-and-spfx-hooks)

Public hook signatures and result shapes match the library source. External types (React, SPFx and PnPjs) come from their respective packages. All hooks require a matching SPFx provider.

## useSPFxThemeInfo

```typescript
export function useSPFxThemeInfo(): IReadonlyTheme | undefined;

```

Returns the current SPFx ThemeProvider theme when available. The provider owns its theme subscription and removes it on unmount.

```tsx
const theme = useSPFxThemeInfo();
return <span>{theme?.isInverted ? 'Inverted host theme' : 'Default or unavailable host theme'}</span>;
```

The result is the native readonly theme reference or undefined before theme availability. It does not create a FluentProvider.

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

Chooses a normalized application SDK theme when context acquisition succeeded and a theme is available, otherwise converts the SPFx theme, otherwise uses the light-theme fallback. SharePoint updates come from the existing provider subscription; Teams/Office/Outlook updates come from the live SDK dispatcher described in [environment hooks](environment.md#usespfxteams).

The legacy `isTeams` field equals SDK context support, including Office/Outlook; it does not classify the actual host or guarantee that an SDK theme is available. Optional `teamsTheme` is the normalized default/dark/highContrast theme. Use `useSPFxEnvironmentInfo` for host classification. There is no `isDark` field. The SDK's single-handler ownership and lack of unregister also apply here.

[WebPartShell](../controls/web-part-shell.md) composes this hook with explicit theme > nearest ancestor Fluent tokens > host tokens, and explicit direction > actual ancestor Fluent direction > SPFx culture. It supplies the owner document but inherits the renderer; hosts must supply a matching Griffel renderer.

```tsx
const { theme, isTeams, teamsTheme } = useSPFxFluent9ThemeInfo();
return <FluentProvider theme={theme}><span>{isTeams ? teamsTheme ?? 'SDK theme unavailable' : 'SharePoint theme'}</span></FluentProvider>;
```

Import `FluentProvider` from `@fluentui/react-components`; the hook returns tokens without mounting a theme boundary.

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

Before an element has been measured, or for a host without a container, `width` and `height` are zero and `size` is `small`. Thus `isSmall` can mean “unmeasured,” rather than a confirmed narrow surface.

```tsx
const { size, width, height, isSmall } = useSPFxContainerSize();
return <p>{size}: {width} × {height}{isSmall ? ' (small or unmeasured)' : ''}</p>;
```

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

- [API index](../INDEX.md)
- [Services](../INDEX.md#non-react-services)
- [SharePoint validation](../../development/sharepoint-validation/INDEX.md)
