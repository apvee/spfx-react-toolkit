# Environment Hooks

> Hooks for environment detection and runtime information

## Overview

These hooks provide access to SPFx runtime environment, Teams integration, locale settings, and page type information.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxEnvironmentInfo`](#usespfxenvironmentinfo) | `SPFxEnvironmentInfo` | Runtime environment details |
| [`useSPFxTeams`](#usespfxteams) | `SPFxTeamsInfo` | Microsoft Teams context |
| [`useSPFxLocaleInfo`](#usespfxlocaleinfo) | `SPFxLocaleInfo` | Locale and culture settings |
| [`useSPFxPageType`](#usespfxpagetype) | `SPFxPageTypeInfo` | Page layout type |

---

## useSPFxEnvironmentInfo

Access SPFx runtime environment information.

### Signature

```typescript
function useSPFxEnvironmentInfo(): SPFxEnvironmentInfo
```

### Returns

```typescript
type EnvironmentType = 'Local' | 'Test' | 'Production';

interface SPFxEnvironmentInfo {
  /** Environment type (Local, Test, or Production) */
  readonly type: EnvironmentType;
  
  /** Whether running in local workbench */
  readonly isLocal: boolean;
  
  /** Whether running in SharePoint Online */
  readonly isSharePointOnline: boolean;
  
  /** Whether running in SharePoint Server (on-premises) */
  readonly isSharePointServer: boolean;
}
```

### Example

```tsx
import { useSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';

function EnvironmentBanner() {
  const { type, isLocal, isSharePointOnline } = useSPFxEnvironmentInfo();
  
  // Show banner in development mode
  if (isLocal) {
    return (
      <div className="dev-banner">
        🛠️ Development Mode - Local Workbench
      </div>
    );
  }
  
  return null;
}
```

### Example: Feature Toggling

```tsx
import { useSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';

function FeatureComponent() {
  const { type, isSharePointOnline } = useSPFxEnvironmentInfo();
  
  // Enable experimental features only in test environment
  const showExperimental = type === 'Test';
  
  // SharePoint Online specific features
  if (!isSharePointOnline) {
    return <p>This feature requires SharePoint Online.</p>;
  }
  
  return (
    <div>
      <h2>Features</h2>
      {showExperimental && (
        <section className="experimental">
          <h3>🧪 Experimental Features</h3>
          {/* ... */}
        </section>
      )}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxEnvironmentInfo.ts)

---

## useSPFxTeams

Access Microsoft Teams context (when running in Teams).

### Signature

```typescript
function useSPFxTeams(): SPFxTeamsInfo
```

### Returns

```typescript
interface TeamsContext {
  /** Team ID (GUID) */
  readonly teamId: string;
  
  /** Team display name */
  readonly teamName: string;
  
  /** Channel ID */
  readonly channelId: string;
  
  /** Channel display name */
  readonly channelName: string;
  
  /** Teams theme (default, dark, contrast) */
  readonly theme: string;
  
  /** User's locale in Teams */
  readonly locale: string;
  
  /** Session ID for correlation */
  readonly sessionId: string;
}

interface SPFxTeamsInfo {
  /** Whether running inside Microsoft Teams */
  readonly isInTeams: boolean;
  
  /** Teams context (only available when isInTeams is true) */
  readonly context: TeamsContext | undefined;
  
  /** Whether running as Teams tab */
  readonly isTab: boolean;
  
  /** Whether running as Teams personal app */
  readonly isPersonalApp: boolean;
}
```

### Example: Teams-Adaptive UI

```tsx
import { useSPFxTeams } from '@apvee/spfx-react-toolkit';

function AdaptiveComponent() {
  const { isInTeams, context, isPersonalApp } = useSPFxTeams();
  
  if (isInTeams && context) {
    return (
      <div className={`teams-container teams-theme-${context.theme}`}>
        {!isPersonalApp && (
          <header>
            <span>Team: {context.teamName}</span>
            <span>Channel: {context.channelName}</span>
          </header>
        )}
        <main>
          {/* Teams-specific UI */}
        </main>
      </div>
    );
  }
  
  return (
    <div className="sharepoint-container">
      {/* SharePoint-specific UI */}
    </div>
  );
}
```

### Example: Teams Deep Linking

```tsx
import { useSPFxTeams } from '@apvee/spfx-react-toolkit';

function DeepLinkButton({ itemId }: { itemId: string }) {
  const { isInTeams, context } = useSPFxTeams();
  
  const handleClick = () => {
    if (isInTeams && context) {
      // Create Teams deep link
      const deepLink = `https://teams.microsoft.com/l/entity/${context.teamId}/` +
        `?context={"subEntityId":"${itemId}"}`;
      
      window.open(deepLink, '_blank');
    }
  };
  
  return (
    <button onClick={handleClick} disabled={!isInTeams}>
      Share in Teams
    </button>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxTeams.ts)

---

## useSPFxLocaleInfo

Access locale and culture settings.

### Signature

```typescript
function useSPFxLocaleInfo(): SPFxLocaleInfo
```

### Returns

```typescript
interface SPFxLocaleInfo {
  /** Current UI culture (e.g., "en-US", "de-DE") */
  readonly currentCulture: string;
  
  /** Current UI language ID (LCID) */
  readonly currentLanguageId: number;
  
  /** Whether UI is right-to-left */
  readonly isRightToLeft: boolean;
  
  /** Site's default language ID */
  readonly siteLanguageId: number;
  
  /** Time zone bias (minutes from UTC) */
  readonly timeZoneBias: number;
}
```

### Example: Locale-Aware Formatting

```tsx
import { useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';

function LocalizedDate({ date }: { date: Date }) {
  const { currentCulture, timeZoneBias } = useSPFxLocaleInfo();
  
  const formatter = new Intl.DateTimeFormat(currentCulture, {
    dateStyle: 'full',
    timeStyle: 'short'
  });
  
  return <time dateTime={date.toISOString()}>{formatter.format(date)}</time>;
}
```

### Example: RTL Support

```tsx
import { useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
import styles from './Component.module.scss';

function LocalizedComponent() {
  const { isRightToLeft, currentCulture } = useSPFxLocaleInfo();
  
  return (
    <div 
      className={styles.container}
      dir={isRightToLeft ? 'rtl' : 'ltr'}
      lang={currentCulture}
    >
      <h2>Content</h2>
      <p>This content respects RTL languages like Arabic and Hebrew.</p>
    </div>
  );
}
```

### Example: Number Formatting

```tsx
import { useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';

function LocalizedCurrency({ amount }: { amount: number }) {
  const { currentCulture } = useSPFxLocaleInfo();
  
  const formatted = new Intl.NumberFormat(currentCulture, {
    style: 'currency',
    currency: 'EUR'
  }).format(amount);
  
  return <span className="price">{formatted}</span>;
}
```

### Source

[View source](../../src/hooks/useSPFxLocaleInfo.ts)

---

## useSPFxPageType

Access page type information.

### Signature

```typescript
function useSPFxPageType(): SPFxPageTypeInfo
```

### Returns

```typescript
type PageType = 
  | 'Unknown'
  | 'WikiPage'
  | 'WebPartPage'
  | 'PublishingPage'
  | 'HomePage'
  | 'SpacesPage'
  | 'SitePage'
  | 'AppPage';

interface SPFxPageTypeInfo {
  /** Numeric page type value */
  readonly pageType: number;
  
  /** Page type as readable string */
  readonly pageTypeName: PageType;
  
  /** Whether this is a modern page */
  readonly isModernPage: boolean;
  
  /** Whether this is the site's home page */
  readonly isHomePage: boolean;
}
```

### Example

```tsx
import { useSPFxPageType } from '@apvee/spfx-react-toolkit';

function PageTypeAwareComponent() {
  const { pageTypeName, isModernPage, isHomePage } = useSPFxPageType();
  
  // Different layout for home page
  if (isHomePage) {
    return <HomePageLayout />;
  }
  
  // Warn about classic pages
  if (!isModernPage) {
    return (
      <div className="warning">
        <p>This web part works best on modern pages.</p>
        <p>Current page type: {pageTypeName}</p>
      </div>
    );
  }
  
  return <StandardLayout pageType={pageTypeName} />;
}
```

### Example: Conditional Features

```tsx
import { useSPFxPageType } from '@apvee/spfx-react-toolkit';

function PublishingFeatures() {
  const { pageTypeName } = useSPFxPageType();
  
  // Only show publishing features on publishing pages
  const isPublishing = pageTypeName === 'PublishingPage';
  
  return (
    <div>
      {isPublishing && (
        <section>
          <h3>Publishing Options</h3>
          {/* Publishing-specific controls */}
        </section>
      )}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxPageType.ts)

---

## See Also

- [Context Hooks](./context.md) - SPFx context access
- [User & Site Hooks](./user-site.md) - User and site information
- [Theming Hooks](./theming.md) - Theme and container information

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
