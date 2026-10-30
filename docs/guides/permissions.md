# Plan permissions and API prechecks

[Documentation home](../INDEX.md)

Choose the check that corresponds to the resource you will access:

| Question | Reference |
| --- | --- |
| Does the user have a permission in the current SharePoint context? | [useSPFxPermissions](../api/hooks/permissions.md#usespfxpermissions) |
| What permissions apply on another web? | [useSPFxCrossSitePermissions](../api/hooks/permissions.md#usespfxcrosssitepermissions) |
| Is the configured delegated Graph/custom API scope present in an acquired token? | [useSPFxApiPermissionPrecheck](../api/hooks/http-clients/api-permission-precheck.md) |
| How can this run outside React? | [Permission helpers](../api/helpers/permission.md), [precheck helpers](../api/helpers/api-permission-precheck.md), [precheck service](../api/services/api-permission-precheck.md) |

A token scope precheck diagnoses configuration and consent; it does not prove that a particular server resource will authorize the request. SharePoint masks, Graph delegated scopes and destination-list grants answer different questions. Handle pending, failed and denied checks separately, then handle actual request failures as documented by the selected client or service.

Follow the precheck reference for remediation entries in `webApiPermissionRequests`, token acquisition behavior, configuration states and retry options. Any required tenant approval remains an explicit administrator operation. See the [authenticated Graph/storage checklist](../development/sharepoint-validation/checklists/graph-storage-permissions.md) for real account and permission evidence.

## Scoped checks and conditional rendering

Use the [scoped permission overload](../api/hooks/permissions.md) for a specific web, list, item or folder. Supply the scope's documented selector; URL/list identity and errors remain distinct from a negative permission result. The legacy staged-URL overload retains its separate execution contract.

[SharePointPermissionTrim](../api/controls/sharepoint-permission-trim.md), [SharePointGroupTrim](../api/controls/sharepoint-group-trim.md) and [EntraGroupTrim](../api/controls/entra-group-trim.md) compose render branches around the current decision. Keep loading, denied and failed states distinguishable. Entra membership checks require the documented existing delegated grants and can be limited by hidden membership or incomplete directory information. Conditional rendering is not server authorization.
