# Permissions Hooks

> Hooks for permission checking and access control

## Overview

These hooks provide permission checking capabilities for the current site and cross-site scenarios.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxPermissions`](#usespfxpermissions) | `SPFxPermissionsResult` | Current site permission checks |
| [`useSPFxCrossSitePermissions`](#usespfxcrosssitepermissions) | `SPFxCrossSitePermissionsResult` | Cross-site permission checks |

---

## useSPFxPermissions

Check user permissions on the current site.

### Signature

```typescript
function useSPFxPermissions(): SPFxPermissionsResult
```

### Returns

```typescript
interface SPFxPermissionsResult {
  /**
   * Check if user has a specific permission.
   * @param permission - SPBasePermissions kind (from @microsoft/sp-page-context)
   * @returns boolean indicating if user has the permission
   */
  readonly hasPermission: (permission: SPBasePermissions) => boolean;
  
  /**
   * Check if user has all specified permissions.
   * @param permissions - Array of SPBasePermissions
   * @returns boolean indicating if user has all permissions
   */
  readonly hasAllPermissions: (permissions: SPBasePermissions[]) => boolean;
  
  /**
   * Check if user has any of the specified permissions.
   * @param permissions - Array of SPBasePermissions
   * @returns boolean indicating if user has at least one permission
   */
  readonly hasAnyPermission: (permissions: SPBasePermissions[]) => boolean;
  
  /** Whether user can add list items */
  readonly canAddListItems: boolean;
  
  /** Whether user can edit list items */
  readonly canEditListItems: boolean;
  
  /** Whether user can delete list items */
  readonly canDeleteListItems: boolean;
  
  /** Whether user can approve items */
  readonly canApproveItems: boolean;
  
  /** Whether user can manage lists */
  readonly canManageLists: boolean;
  
  /** Whether user can manage web */
  readonly canManageWeb: boolean;
  
  /** Whether user can view pages */
  readonly canViewPages: boolean;
  
  /** Whether user is site collection admin */
  readonly isSiteAdmin: boolean;
}
```

### Example: Conditional UI

```tsx
import { useSPFxPermissions } from '@apvee/spfx-react-toolkit';

function ItemActions({ itemId }: { itemId: number }) {
  const { canEditListItems, canDeleteListItems, canApproveItems } = useSPFxPermissions();
  
  return (
    <div className="actions">
      {canEditListItems && (
        <button onClick={() => handleEdit(itemId)}>Edit</button>
      )}
      
      {canDeleteListItems && (
        <button onClick={() => handleDelete(itemId)}>Delete</button>
      )}
      
      {canApproveItems && (
        <button onClick={() => handleApprove(itemId)}>Approve</button>
      )}
    </div>
  );
}
```

### Example: Admin Panel

```tsx
import { useSPFxPermissions } from '@apvee/spfx-react-toolkit';

function AdminPanel() {
  const { isSiteAdmin, canManageWeb, canManageLists } = useSPFxPermissions();
  
  if (!isSiteAdmin && !canManageWeb) {
    return (
      <MessageBar messageBarType={MessageBarType.blocked}>
        You don't have permission to access admin settings.
      </MessageBar>
    );
  }
  
  return (
    <div className="admin-panel">
      <h2>Administration</h2>
      
      <section>
        <h3>Site Settings</h3>
        {/* Site settings available to site admins and web managers */}
      </section>
      
      {canManageLists && (
        <section>
          <h3>List Management</h3>
          {/* List management options */}
        </section>
      )}
    </div>
  );
}
```

### Example: Custom Permission Check

```tsx
import { useSPFxPermissions } from '@apvee/spfx-react-toolkit';
import { SPBasePermissions } from '@microsoft/sp-page-context';

function ContentApprovalWorkflow() {
  const { hasPermission, hasAllPermissions, hasAnyPermission } = useSPFxPermissions();
  
  // Check single permission
  const canCreateFolders = hasPermission(SPBasePermissions.CreateGroups);
  
  // Check multiple permissions (ALL required)
  const canManageContent = hasAllPermissions([
    SPBasePermissions.ManageLists,
    SPBasePermissions.EditListItems,
    SPBasePermissions.ApproveItems
  ]);
  
  // Check multiple permissions (ANY sufficient)
  const canContribute = hasAnyPermission([
    SPBasePermissions.AddListItems,
    SPBasePermissions.EditListItems
  ]);
  
  return (
    <div className="workflow">
      {canContribute && <button>Submit for Review</button>}
      {canManageContent && <button>Publish</button>}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxPermissions.ts)

---

## useSPFxCrossSitePermissions

Check user permissions on a different site.

### Signature

```typescript
function useSPFxCrossSitePermissions(siteUrl: string): SPFxCrossSitePermissionsResult
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `siteUrl` | `string` | Yes | Absolute URL of the target site |

### Returns

```typescript
interface SPFxCrossSitePermissionsResult {
  /** Loading state */
  readonly isLoading: boolean;
  
  /** Error if permission check failed */
  readonly error: Error | undefined;
  
  /**
   * Check if user has a specific permission.
   * Returns undefined while loading.
   */
  readonly hasPermission: (permission: SPBasePermissions) => boolean | undefined;
  
  /**
   * Check if user has all specified permissions.
   * Returns undefined while loading.
   */
  readonly hasAllPermissions: (permissions: SPBasePermissions[]) => boolean | undefined;
  
  /**
   * Check if user has any of the specified permissions.
   * Returns undefined while loading.
   */
  readonly hasAnyPermission: (permissions: SPBasePermissions[]) => boolean | undefined;
  
  /** Whether user can add list items (undefined while loading) */
  readonly canAddListItems: boolean | undefined;
  
  /** Whether user can edit list items (undefined while loading) */
  readonly canEditListItems: boolean | undefined;
  
  /** Whether user can delete list items (undefined while loading) */
  readonly canDeleteListItems: boolean | undefined;
  
  /** Whether user can manage lists (undefined while loading) */
  readonly canManageLists: boolean | undefined;
  
  /** Refresh permissions */
  readonly refresh: () => void;
}
```

### Example: Cross-Site Item Copy

```tsx
import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';

function CopyToSite({ item, targetSiteUrl }: { item: IItem; targetSiteUrl: string }) {
  const { canAddListItems, isLoading, error } = useSPFxCrossSitePermissions(targetSiteUrl);
  
  if (isLoading) {
    return <Spinner label="Checking permissions..." />;
  }
  
  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error.message}</MessageBar>;
  }
  
  if (!canAddListItems) {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        You don't have permission to add items to the target site.
      </MessageBar>
    );
  }
  
  return (
    <button onClick={() => copyItem(item, targetSiteUrl)}>
      Copy to {targetSiteUrl}
    </button>
  );
}
```

### Example: Multi-Site Dashboard

```tsx
import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';

interface SitePermissionStatus {
  url: string;
  name: string;
}

function SitePermissionCard({ site }: { site: SitePermissionStatus }) {
  const { 
    canEditListItems, 
    canManageLists, 
    isLoading, 
    error,
    refresh 
  } = useSPFxCrossSitePermissions(site.url);
  
  return (
    <div className="site-card">
      <h3>{site.name}</h3>
      <p>{site.url}</p>
      
      {isLoading ? (
        <Spinner size={SpinnerSize.small} />
      ) : error ? (
        <span className="error">Unable to check permissions</span>
      ) : (
        <div className="permissions">
          <span className={canEditListItems ? 'granted' : 'denied'}>
            {canEditListItems ? '✓' : '✗'} Edit
          </span>
          <span className={canManageLists ? 'granted' : 'denied'}>
            {canManageLists ? '✓' : '✗'} Manage
          </span>
        </div>
      )}
      
      <button onClick={refresh}>Refresh</button>
    </div>
  );
}

function MultiSiteDashboard({ sites }: { sites: SitePermissionStatus[] }) {
  return (
    <div className="site-grid">
      {sites.map(site => (
        <SitePermissionCard key={site.url} site={site} />
      ))}
    </div>
  );
}
```

### Example: Hub Site Permissions

```tsx
import { 
  useSPFxCrossSitePermissions, 
  useSPFxHubSiteInfo,
  useSPFxSiteInfo 
} from '@apvee/spfx-react-toolkit';

function HubSiteAdminCheck() {
  const { isConnected, hubSiteId } = useSPFxHubSiteInfo();
  const { siteUrl } = useSPFxSiteInfo();
  
  // Only check hub permissions if connected to a hub
  const hubUrl = isConnected ? `https://contoso.sharepoint.com/sites/hub-${hubSiteId}` : '';
  const hubPermissions = useSPFxCrossSitePermissions(hubUrl);
  
  if (!isConnected) {
    return <p>Not connected to a hub site.</p>;
  }
  
  if (hubPermissions.isLoading) {
    return <Spinner label="Checking hub site permissions..." />;
  }
  
  const isHubAdmin = hubPermissions.canManageLists;
  
  return (
    <div>
      {isHubAdmin ? (
        <span className="badge admin">Hub Administrator</span>
      ) : (
        <span className="badge member">Hub Member</span>
      )}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxCrossSitePermissions.ts)

---

## Permission Constants Reference

Common `SPBasePermissions` values from `@microsoft/sp-page-context`:

| Permission | Description |
|------------|-------------|
| `ViewListItems` | View items in lists |
| `AddListItems` | Add items to lists |
| `EditListItems` | Edit items in lists |
| `DeleteListItems` | Delete items from lists |
| `ApproveItems` | Approve content |
| `OpenItems` | Open items |
| `ViewVersions` | View item versions |
| `DeleteVersions` | Delete item versions |
| `ManageLists` | Create/delete lists |
| `ManageWeb` | Manage site settings |
| `AddAndCustomizePages` | Add/edit pages |
| `BrowseDirectories` | Browse directories |
| `CreateGroups` | Create SharePoint groups |
| `ManagePermissions` | Manage permissions |
| `FullMask` | Full control |

---

## See Also

- [User & Site Hooks](./user-site.md) - User information
- [Context Hooks](./context.md) - SPFx context access
- [Environment Hooks](./environment.md) - Environment detection

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
