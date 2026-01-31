# User & Site Hooks

> Hooks for accessing user and site information

## Overview

These hooks provide access to current user, site, hub site, and list information.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxUserInfo`](#usespfxuserinfo) | `SPFxUserInfo` | Current user information |
| [`useSPFxUserPhoto`](#usespfxuserphoto) | `SPFxUserPhotoResult` | User profile photo |
| [`useSPFxSiteInfo`](#usespfxsiteinfo) | `SPFxSiteInfo` | Site and web information |
| [`useSPFxHubSiteInfo`](#usespfxhubsiteinfo) | `SPFxHubSiteInfo` | Hub site information |
| [`useSPFxListInfo`](#usespfxlistinfo) | `SPFxListInfo \| undefined` | Current list context |

---

## useSPFxUserInfo

Access current user information.

### Signature

```typescript
function useSPFxUserInfo(): SPFxUserInfo
```

### Returns

```typescript
interface SPFxUserInfo {
  /** User login name (e.g., "domain\\user" or email) */
  readonly loginName: string;
  
  /** User display name */
  readonly displayName: string;
  
  /** User email address (optional) */
  readonly email?: string;
  
  /** Whether user is an external guest user */
  readonly isExternal: boolean;
}
```

### Example

```tsx
import { useSPFxUserInfo } from '@apvee/spfx-react-toolkit';

function WelcomeMessage() {
  const { displayName, email, isExternal } = useSPFxUserInfo();
  
  return (
    <div>
      <h2>Welcome, {displayName}!</h2>
      {email && <p>Email: {email}</p>}
      {isExternal && (
        <span className="badge">Guest User</span>
      )}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxUserInfo.ts)

---

## useSPFxUserPhoto

Access user profile photo with multiple size options.

### Signature

```typescript
function useSPFxUserPhoto(options?: SPFxUserPhotoOptions): SPFxUserPhotoResult
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `options` | `SPFxUserPhotoOptions` | No | Photo options |

### Options

```typescript
type SPFxUserPhotoSize = 
  | 'S'   // 48×48
  | 'M'   // 72×72  
  | 'L';  // 120×120

interface SPFxUserPhotoOptions {
  /** User login name (default: current user) */
  loginName?: string;
  
  /** Photo size (default: 'M') */
  size?: SPFxUserPhotoSize;
}
```

### Returns

```typescript
interface SPFxUserPhotoResult {
  /** Photo URL (undefined while loading) */
  readonly url: string | undefined;
  
  /** Loading state */
  readonly isLoading: boolean;
  
  /** Error if failed */
  readonly error: Error | undefined;
}
```

### Example: Current User Photo

```tsx
import { useSPFxUserPhoto, useSPFxUserInfo } from '@apvee/spfx-react-toolkit';

function UserCard() {
  const { displayName, email } = useSPFxUserInfo();
  const { url, isLoading } = useSPFxUserPhoto({ size: 'L' });
  
  return (
    <div className="user-card">
      {isLoading ? (
        <div className="avatar-placeholder" />
      ) : (
        <img src={url} alt={displayName} className="avatar" />
      )}
      <h3>{displayName}</h3>
      <p>{email}</p>
    </div>
  );
}
```

### Example: Other User Photo

```tsx
import { useSPFxUserPhoto } from '@apvee/spfx-react-toolkit';

function TeamMemberCard({ loginName, name }: { loginName: string; name: string }) {
  const { url, isLoading, error } = useSPFxUserPhoto({ 
    loginName, 
    size: 'M' 
  });
  
  return (
    <div className="team-member">
      <img 
        src={error ? '/images/default-avatar.png' : url} 
        alt={name}
      />
      <span>{name}</span>
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxUserPhoto.ts)

---

## useSPFxSiteInfo

Access site collection and web information.

### Signature

```typescript
function useSPFxSiteInfo(): SPFxSiteInfo
```

### Returns

```typescript
interface SPFxGroupInfo {
  /** Group ID (GUID) */
  readonly id: string;
  
  /** Whether group is public */
  readonly isPublic: boolean;
}

interface SPFxSiteInfo {
  // Web properties
  /** Web ID (GUID) */
  readonly webId: string;
  
  /** Web absolute URL */
  readonly webUrl: string;
  
  /** Web server relative URL */
  readonly webServerRelativeUrl: string;
  
  /** Web title */
  readonly title: string;
  
  /** Web language ID (LCID) */
  readonly languageId: number;
  
  /** Site logo URL */
  readonly logoUrl?: string;
  
  // Site collection properties
  /** Site collection ID (GUID) */
  readonly siteId: string;
  
  /** Site collection absolute URL */
  readonly siteUrl: string;
  
  /** Site collection server relative URL */
  readonly siteServerRelativeUrl: string;
  
  /** Site classification label */
  readonly siteClassification?: string;
  
  /** Microsoft 365 Group info (if group-connected) */
  readonly siteGroup?: SPFxGroupInfo;
}
```

### Example

```tsx
import { useSPFxSiteInfo } from '@apvee/spfx-react-toolkit';

function SiteHeader() {
  const { 
    title, 
    logoUrl, 
    webUrl,
    siteClassification, 
    siteGroup 
  } = useSPFxSiteInfo();
  
  return (
    <header>
      {logoUrl && <img src={logoUrl} alt="Site logo" />}
      <h1>{title}</h1>
      
      {siteClassification && (
        <span className="classification">{siteClassification}</span>
      )}
      
      {siteGroup && (
        <span className="group-badge">
          {siteGroup.isPublic ? 'Public Group' : 'Private Group'}
        </span>
      )}
    </header>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxSiteInfo.ts)

---

## useSPFxHubSiteInfo

Access hub site information.

### Signature

```typescript
function useSPFxHubSiteInfo(): SPFxHubSiteInfo
```

### Returns

```typescript
interface SPFxHubSiteInfo {
  /** Whether current site is connected to a hub */
  readonly isConnected: boolean;
  
  /** Hub site ID (GUID) */
  readonly hubSiteId?: string;
  
  /** Whether current site is a hub site */
  readonly isHubSite: boolean;
}
```

### Example

```tsx
import { useSPFxHubSiteInfo, useSPFxSiteInfo } from '@apvee/spfx-react-toolkit';

function HubNavigation() {
  const { isConnected, isHubSite, hubSiteId } = useSPFxHubSiteInfo();
  const { title } = useSPFxSiteInfo();
  
  if (!isConnected && !isHubSite) {
    return null; // No hub connection
  }
  
  return (
    <nav className="hub-nav">
      {isHubSite ? (
        <span className="hub-badge">Hub Site</span>
      ) : (
        <span>Connected to Hub: {hubSiteId}</span>
      )}
    </nav>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxHubSiteInfo.ts)

---

## useSPFxListInfo

Access current list context (when in list context).

### Signature

```typescript
function useSPFxListInfo(): SPFxListInfo | undefined
```

### Returns

```typescript
interface SPFxListInfo {
  /** List ID (GUID) */
  readonly id: string;
  
  /** List title */
  readonly title: string;
  
  /** List server relative URL */
  readonly serverRelativeUrl: string;
  
  /** List base template ID */
  readonly baseTemplate: number;
}
```

Returns `undefined` if not in a list context (e.g., on a site page without list association).

### Example

```tsx
import { useSPFxListInfo } from '@apvee/spfx-react-toolkit';

function ListAwareComponent() {
  const listInfo = useSPFxListInfo();
  
  if (!listInfo) {
    return <p>This component requires a list context.</p>;
  }
  
  return (
    <div>
      <h2>List: {listInfo.title}</h2>
      <p>Template: {listInfo.baseTemplate}</p>
      <p>URL: {listInfo.serverRelativeUrl}</p>
    </div>
  );
}
```

### Example: Field Customizer Usage

```tsx
import { useSPFxListInfo } from '@apvee/spfx-react-toolkit';

function FieldRenderer({ value }: { value: string }) {
  const listInfo = useSPFxListInfo();
  
  // Field customizer always has list context
  const isTaskList = listInfo?.baseTemplate === 107; // Task list template
  
  return (
    <span className={isTaskList ? 'task-field' : 'generic-field'}>
      {value}
    </span>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxListInfo.ts)

---

## See Also

- [Permissions Hooks](./permissions.md) - Permission checking
- [Environment Hooks](./environment.md) - Environment detection
- [Context Hooks](./context.md) - Context access

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
