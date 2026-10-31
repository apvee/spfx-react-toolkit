# Access trims and scoped permissions

[Validation index](../INDEX.md) · [Setup](../setup.md)

These authenticated checks remain NOT EXECUTED until recorded against an actual tenant/user. Local deterministic HTTP/Graph fixtures and builds are not tenant passes.

- In each actual lazy trim panel, Run/repeat a permitted, denied and invalid target. Observe Loading, protected counter/effect mounts, distinct Denied/Error text and the applied target. Change target/user/client/provider and confirm former content disappears before replacement reads complete.
- SharePoint Permission Trim: current-web PageContext mask and native fallback, remote web, list by ID/title/decoded path, item, folder with parent-list validation; library root folder without list item uses list scope. Verify all/any and nested permission AND Entra group composition. Confirm repeat starts a new read where native mask is absent.
- SharePoint Group Trim: direct membership, nonmembership, remote site-local user identity, exact names including apostrophes, all pages, missing/ambiguous groups and unreadable group errors. Entra principals within site groups are not expanded.
- Entra Group Trim: obtain tenant approval for the configured User.Read and GroupMember.Read.All requests through the normal process; verify actual Graph access separately. Test GUID/name, All/M365/Security, member/owner/memberOrOwner, all/any, more than twenty IDs, type mismatch, name filtering before ambiguity and owner paging with limited-information IDs. Synchronized/distribution/non-M365 mail-enabled owners must report Error. Verify hidden membership limits; omissions are not proof of absence.
- Context: preserve legacy idle/staged URL behavior and separately Run/repeat all four scoped overload targets; report errors independently from denial. Confirm actual server operations still enforce authorization.
- Record tenant, host (SharePoint/Teams), user permissions, consent status, dates, expected/observed results and limitations in maintenance evidence. Record changes and cleanup for the actual run.
