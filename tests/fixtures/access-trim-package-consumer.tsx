import * as React from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit/controls/sharepoint-permission-trim';
import type { SharePointPermissionTrimProps, SharePointPermissionTrimBranchProps } from '@apvee/spfx-react-toolkit/controls/sharepoint-permission-trim';
import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit/controls/sharepoint-group-trim';
import type { SharePointGroupSelector, SharePointGroupTrimProps, SharePointGroupTrimBranchProps } from '@apvee/spfx-react-toolkit/controls/sharepoint-group-trim';
import { EntraGroupTrim } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';
import type { EntraGroupSelector, EntraGroupTrimProps, EntraGroupTrimBranchProps, EntraGroupType, EntraGroupRelationship } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';

const permission: SharePointPermissionTrimProps = { permissions: [SPPermission.addListItems], scope: 'item', listPath: 'Lists/Tasks', itemId: 42 };
const permissionBranch: SharePointPermissionTrimBranchProps = { children: <span>Checking</span> };
<SharePointPermissionTrim {...permission}><SharePointPermissionTrim.Loading {...permissionBranch} /></SharePointPermissionTrim>;
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} />;
// @ts-expect-error scope list requires a selector
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} scope="list" />;
// @ts-expect-error item fields belong only to item scope
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} scope="web" itemId={1} />;
const bothLists = { scope: 'list' as const, listId: 'a', listTitle: 'Tasks', permissions: [SPPermission.viewListItems] };
// @ts-expect-error list selector XOR also applies to spreads
<SharePointPermissionTrim {...bothLists} />;
// @ts-expect-error match is exactly all/any
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} match="some" />;
// @ts-expect-error no style forwarding on headless control
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} style={{ color: 'red' }} />;

const spSelector: SharePointGroupSelector = { id: 2 };
const spProps: SharePointGroupTrimProps = { groups: [spSelector], webUrl: 'https://tenant.test/sites/current', match: 'any' };
const spBranch: SharePointGroupTrimBranchProps = { children: 'No membership' };
<SharePointGroupTrim {...spProps}><SharePointGroupTrim.Fallback {...spBranch} /></SharePointGroupTrim>;
// @ts-expect-error group selector requires ID XOR name
const spBoth: SharePointGroupSelector = { id: 2, name: 'Readers' };
const spSpread = { id: 2, name: 'Readers' };
// @ts-expect-error XOR survives spread object assignment
<SharePointGroupTrim groups={[{ ...spSpread }]} />;
// @ts-expect-error no numeric string IDs
<SharePointGroupTrim groups={[{ id: '2' }]} />;
// @ts-expect-error branches accept children only
<SharePointGroupTrim.Loading className="loading" />;

const type: EntraGroupType = 'Security';
const relationship: EntraGroupRelationship = 'memberOrOwner';
const selector: EntraGroupSelector = { id: 'aaaaaaaa-0000-0000-0000-000000000002' };
const entraProps: EntraGroupTrimProps = { groups: [selector], groupType: type, relationship, match: 'any' };
const entraBranch: EntraGroupTrimBranchProps = { children: 'Unavailable' };
<EntraGroupTrim {...entraProps}><EntraGroupTrim.ErrorFallback {...entraBranch} /></EntraGroupTrim>;
// @ts-expect-error group selector ID/name XOR
const entraBoth: EntraGroupSelector = { id: 'a', name: 'Team' };
// @ts-expect-error type belongs only on control
<EntraGroupTrim groups={[{ id: 'a', groupType: 'M365' }]} />;
const typedSelector = { id: 'a', groupType: 'M365' as const };
// @ts-expect-error selector groupType rejection survives spread
<EntraGroupTrim groups={[{ ...typedSelector }]} />;
const namedType = { name: 'Team', groupType: 'Security' as const };
// @ts-expect-error named selector groupType also survives assignment
const forbidden: EntraGroupSelector = namedType;
// @ts-expect-error groupType is case-sensitive
<EntraGroupTrim groups={[selector]} groupType="security" />;
// @ts-expect-error relationship constrained
<EntraGroupTrim groups={[selector]} relationship="owners" />;
// @ts-expect-error no asynchronous public result callbacks
<EntraGroupTrim groups={[selector]} onCheck={() => 'No'} />;
// ReactNode interoperability retains legacy consumer null only.
<EntraGroupTrim groups={[selector]}>{null}<EntraGroupTrim.Loading /></EntraGroupTrim>;
void [spBoth, entraBoth, forbidden];

import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit/hooks';
import type { SPFxScopedPermissionsOptions, SPFxScopedPermissionsInfo, SPFxCrossSitePermissionsInfo } from '@apvee/spfx-react-toolkit';
const web: SPFxScopedPermissionsOptions = {};
const list: SPFxScopedPermissionsOptions = { scope: 'list', listId: '11111111-2222-3333-4444-555555555555' };
const item: SPFxScopedPermissionsOptions = { scope: 'item', listTitle: 'Tasks', itemId: 42 };
const folder: SPFxScopedPermissionsOptions = { scope: 'folder', listPath: 'Shared Documents', folderPath: '/sites/team/Shared Documents/Folder' };
const scoped: SPFxScopedPermissionsInfo = useSPFxCrossSitePermissions(item);
const legacy: SPFxCrossSitePermissionsInfo = useSPFxCrossSitePermissions(undefined, { listId: 'guid' });
scoped.hasPermission(SPPermission.editListItems);
legacy.hasWebPermission(SPPermission.viewListItems);
void [web, list, folder];
// @ts-expect-error List scope requires a selector.
const missing: SPFxScopedPermissionsOptions = { scope: 'list' };
// @ts-expect-error List selectors are mutually exclusive.
const multiple: SPFxScopedPermissionsOptions = { scope: 'list', listId: 'guid', listTitle: 'Tasks' };
// @ts-expect-error Item scope requires its item ID.
const missingItem: SPFxScopedPermissionsOptions = { scope: 'item', listTitle: 'Tasks' };
// @ts-expect-error Item IDs are accepted only for item scope.
const extraItem: SPFxScopedPermissionsOptions = { scope: 'list', listTitle: 'Tasks', itemId: 42 };
// @ts-expect-error Web scope cannot select a list.
const extraList: SPFxScopedPermissionsOptions = { listTitle: 'Tasks' };
// @ts-expect-error Folder scope requires its path.
const missingFolder: SPFxScopedPermissionsOptions = { scope: 'folder', listTitle: 'Docs' };
// @ts-expect-error Folder path is accepted only for folder scope.
const extraFolder: SPFxScopedPermissionsOptions = { scope: 'item', listTitle: 'Tasks', itemId: 1, folderPath: '/Docs' };
void [missing, multiple, missingItem, extraItem, extraList, missingFolder, extraFolder];

<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} scope="list" listTitle="Tasks">
  Protected
  <SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading>
  <SharePointPermissionTrim.Fallback>Denied</SharePointPermissionTrim.Fallback>
  <SharePointPermissionTrim.ErrorFallback>Error</SharePointPermissionTrim.ErrorFallback>
</SharePointPermissionTrim>;
<SharePointGroupTrim groups={[{ name: 'Readers' }]}>
  Protected<SharePointGroupTrim.Loading>Checking</SharePointGroupTrim.Loading>
  <SharePointGroupTrim.Fallback>Denied</SharePointGroupTrim.Fallback>
  <SharePointGroupTrim.ErrorFallback>Error</SharePointGroupTrim.ErrorFallback>
</SharePointGroupTrim>;
<EntraGroupTrim groups={[{ name: 'Team' }]} groupType="All" relationship="member">
  Protected<EntraGroupTrim.Loading>Checking</EntraGroupTrim.Loading>
  <EntraGroupTrim.Fallback>Denied</EntraGroupTrim.Fallback>
  <EntraGroupTrim.ErrorFallback>Error</EntraGroupTrim.ErrorFallback>
</EntraGroupTrim>;
<EntraGroupTrim groups={[selector]} groupType="M365" relationship="owner" />;
<EntraGroupTrim groups={[selector]} groupType="Security" relationship="memberOrOwner" />;
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} scope="folder" listId="11111111-2222-3333-4444-555555555555" folderPath="/sites/team/Docs/Folder" />;
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]} scope="list" listPath="Lists/Tasks" />;
useSPFxCrossSitePermissions();
useSPFxCrossSitePermissions('https://tenant.test/sites/team', { webUrl: 'https://tenant.test/sites/team', listId: '11111111-2222-3333-4444-555555555555' });
