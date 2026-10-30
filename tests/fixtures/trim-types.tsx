import * as React from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { SharePointPermissionTrim } from '../../packages/spfx-react-toolkit/src/controls/sharepoint-permission-trim/sharepoint-permission-trim';
import type { SharePointPermissionTrimProps, SharePointPermissionTrimBranchProps } from '../../packages/spfx-react-toolkit/src/controls/sharepoint-permission-trim/sharepoint-permission-trim.types';
import { SharePointGroupTrim } from '../../packages/spfx-react-toolkit/src/controls/sharepoint-group-trim/sharepoint-group-trim';
import type { SharePointGroupSelector, SharePointGroupTrimProps, SharePointGroupTrimBranchProps } from '../../packages/spfx-react-toolkit/src/controls/sharepoint-group-trim/sharepoint-group-trim.types';
import { EntraGroupTrim } from '../../packages/spfx-react-toolkit/src/controls/entra-group-trim/entra-group-trim';
import type { EntraGroupSelector, EntraGroupTrimProps, EntraGroupTrimBranchProps, EntraGroupType, EntraGroupRelationship } from '../../packages/spfx-react-toolkit/src/controls/entra-group-trim/entra-group-trim.types';

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
