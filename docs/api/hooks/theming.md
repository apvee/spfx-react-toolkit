# UI & Theming Hooks

> Hooks for accessing theme and layout information

## Overview

These hooks provide access to SPFx themes and container sizing for responsive design.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxThemeInfo`](#usespfxthemeinfo) | `IReadonlyTheme \| undefined` | SPFx theme (Fluent UI 8) |
| [`useSPFxFluent9ThemeInfo`](#usespfxfluent9themeinfo) | `SPFxFluent9ThemeInfo` | Fluent UI 9 theme |
| [`useSPFxContainerSize`](#usespfxcontainersize) | `SPFxContainerSizeInfo` | Responsive breakpoints |
| [`useSPFxContainerInfo`](#usespfxcontainerinfo) | `SPFxContainerInfo` | Container dimensions |

---

## useSPFxThemeInfo

Access the current SPFx theme (Fluent UI 8 format).

### Signature

```typescript
function useSPFxThemeInfo(): IReadonlyTheme | undefined
```

### Returns

`IReadonlyTheme | undefined` - SharePoint theme object from `@microsoft/sp-component-base`

### Description

Theme subscription is managed automatically by SPFxProvider. Updates when user switches between light/dark theme or theme settings change.

**Theme object includes:**
- `semanticColors` - Context-aware colors (bodyBackground, bodyText, link, etc.)
- `palette` - Full color palette (themePrimary, neutralLight, etc.)
- `fonts` - Font styles (small, medium, large, etc.)
- `isInverted` - Whether theme is dark mode

### Example: Basic Theme Usage

```tsx
import { useSPFxThemeInfo } from '@apvee/spfx-react-toolkit';

function ThemedComponent() {
  const theme = useSPFxThemeInfo();
  
  return (
    <div style={{ 
      backgroundColor: theme?.semanticColors?.bodyBackground,
      color: theme?.semanticColors?.bodyText,
      padding: '16px',
      borderRadius: '4px',
      border: `1px solid ${theme?.semanticColors?.bodyDivider}`
    }}>
      <h2 style={{ color: theme?.palette?.themePrimary }}>
        Themed Content
      </h2>
      <p>This component respects the SharePoint theme.</p>
    </div>
  );
}
```

### Example: Dark Mode Detection

```tsx
import { useSPFxThemeInfo } from '@apvee/spfx-react-toolkit';

function AdaptiveComponent() {
  const theme = useSPFxThemeInfo();
  const isDarkMode = theme?.isInverted ?? false;
  
  return (
    <div className={isDarkMode ? 'dark-mode' : 'light-mode'}>
      <img 
        src={isDarkMode ? '/images/logo-light.png' : '/images/logo-dark.png'} 
        alt="Logo" 
      />
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxThemeInfo.ts)

---

## useSPFxFluent9ThemeInfo

Access theme converted to Fluent UI 9 format.

### Signature

```typescript
function useSPFxFluent9ThemeInfo(): SPFxFluent9ThemeInfo
```

### Returns

```typescript
interface SPFxFluent9ThemeInfo {
  /** Fluent UI 9 compatible theme object */
  readonly theme: Theme;
  
  /** Whether current theme is dark mode */
  readonly isDark: boolean;
}
```

### Description

Automatically converts SPFx theme to Fluent UI 9 format using `@fluentui/react-migration-v8-v9`. Useful when using Fluent UI 9 components in SPFx.

### Example: With Fluent UI 9

```tsx
import { useSPFxFluent9ThemeInfo } from '@apvee/spfx-react-toolkit';
import { FluentProvider, Button, Card, Text } from '@fluentui/react-components';

function ModernComponent() {
  const { theme, isDark } = useSPFxFluent9ThemeInfo();
  
  return (
    <FluentProvider theme={theme}>
      <Card>
        <Text>This uses Fluent UI 9 with SharePoint theme!</Text>
        <Button appearance="primary">
          {isDark ? '🌙 Dark Mode' : '☀️ Light Mode'}
        </Button>
      </Card>
    </FluentProvider>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxFluent9ThemeInfo.ts)

---

## useSPFxContainerSize

Access responsive container size breakpoints.

### Signature

```typescript
function useSPFxContainerSize(): SPFxContainerSizeInfo
```

### Returns

```typescript
type SPFxContainerSize = 
  | 'tiny'     // < 480px
  | 'small'    // 480-639px
  | 'medium'   // 640-1023px
  | 'large'    // 1024-1365px
  | 'xlarge';  // >= 1366px

interface SPFxContainerSizeInfo {
  /** Current size breakpoint */
  readonly size: SPFxContainerSize;
  
  /** Current width in pixels */
  readonly width: number;
  
  /** Current height in pixels */
  readonly height: number;
  
  /** Is 'tiny' or smaller */
  readonly isTiny: boolean;
  
  /** Is 'small' or smaller */
  readonly isSmall: boolean;
  
  /** Is 'medium' or smaller */
  readonly isMedium: boolean;
  
  /** Is 'large' or smaller */
  readonly isLarge: boolean;
  
  /** Is 'xlarge' */
  readonly isXLarge: boolean;
}
```

### Description

Monitors the SPFx container width and provides semantic breakpoints for responsive design. Updates automatically when container is resized.

**Breakpoint ranges:**
| Size | Width Range |
|------|-------------|
| `tiny` | < 480px |
| `small` | 480-639px |
| `medium` | 640-1023px |
| `large` | 1024-1365px |
| `xlarge` | ≥ 1366px |

### Example: Responsive Layout

```tsx
import { useSPFxContainerSize } from '@apvee/spfx-react-toolkit';

function ResponsiveGrid() {
  const { size, isTiny, isSmall } = useSPFxContainerSize();
  
  // Determine grid columns based on size
  const columns = isTiny ? 1 : isSmall ? 2 : 3;
  
  return (
    <div>
      <p>Current size: {size}</p>
      <div style={{ 
        display: 'grid', 
        gridTemplateColumns: `repeat(${columns}, 1fr)`,
        gap: '16px'
      }}>
        {items.map(item => <Card key={item.id} item={item} />)}
      </div>
    </div>
  );
}
```

### Example: Component Variants

```tsx
import { useSPFxContainerSize } from '@apvee/spfx-react-toolkit';

function AdaptiveNavigation() {
  const { isTiny, isSmall } = useSPFxContainerSize();
  
  // Mobile: hamburger menu
  if (isTiny || isSmall) {
    return <MobileNavigation />;
  }
  
  // Desktop: full navigation bar
  return <DesktopNavigation />;
}
```

### Source

[View source](../../src/hooks/useSPFxContainerSize.ts)

---

## useSPFxContainerInfo

Access raw container dimensions.

### Signature

```typescript
function useSPFxContainerInfo(): SPFxContainerInfo
```

### Returns

```typescript
interface SPFxContainerInfo {
  /** Container width in pixels */
  readonly width: number;
  
  /** Container height in pixels */
  readonly height: number;
}
```

### Description

Provides raw container dimensions in pixels. For semantic breakpoints, use `useSPFxContainerSize` instead.

### Example

```tsx
import { useSPFxContainerInfo } from '@apvee/spfx-react-toolkit';

function DimensionsDisplay() {
  const { width, height } = useSPFxContainerInfo();
  
  return (
    <div>
      <p>Container: {width}×{height}px</p>
      <div style={{
        width: '100%',
        height: Math.min(height * 0.5, 400),
        background: 'lightblue'
      }}>
        Dynamic height content
      </div>
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxContainerInfo.ts)

---

## See Also

- [Environment Hooks](./environment.md) - Environment detection
- [Properties Hooks](./properties.md) - Display mode
- [Context Hooks](./context.md) - Context access

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
