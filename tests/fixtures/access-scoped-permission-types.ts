import { SPPermission } from '@microsoft/sp-page-context';
import { useSPFxCrossSitePermissions } from '../../packages/spfx-react-toolkit/src/hooks/useSPFxCrossSitePermissions';
import type { SPFxScopedPermissionsOptions, SPFxScopedPermissionsInfo, SPFxCrossSitePermissionsInfo } from '../../packages/spfx-react-toolkit/src/hooks/useSPFxCrossSitePermissions';
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
