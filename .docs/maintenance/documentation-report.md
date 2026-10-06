# F7 — Public documentation alignment

Author: documentation_alignment. Date: 2026-10-06. Worktree: `/Users/fabiofranzini/.codex/worktrees/spfx-toolkit-monorepo/spfx-react-toolkit`.

## Result

**Local documentation checks PASS. Independent review remains required. Authenticated tenant checks NOT EXECUTED.** No commit, push, merge, publication or deployment was performed by this task. The only source change is the package-name typo in `packages/spfx-react-toolkit/src/index.ts` JSDoc (`spfx-react-toolkit` → `@apvee/spfx-react-toolkit`); runtime implementation, signatures and exports were not changed.

## Changes and evidence basis

- Root README and npm-package README now distinguish the publishable TypeScript ESNext library workspace from the private SPFx 1.21.1 Gulp sample. Commands were checked against the current root and both workspace manifests. Node range is `>=22.14.0 <23.0.0`; verified repository baseline is React/ReactDOM 17.0.1, SPFx 1.21.1, TypeScript 5.3.3 and PnPjs 4.17.0. Wider SPFx peers remain a declaration, not a tested matrix.
- `docs/DEVELOPMENT.md` records the workspace layout, actual script tasks, app certificate/serve commands, rebuild-library + restart-serve requirement, migration without package/API change, tarball file groups and limits of local verification. npm installation is not described as proof every dependency lifecycle script ran.
- `docs/SHAREPOINT-VALIDATION.md` provides explicit test setup, app serve URL/config/extension ID, actual WebPart/Application Customizer entry points, two-instance isolation, pane-after-setter, mode/theme/resize, placeholder disposal/recreation/unmount, async error/loading, CRUD/paging/refiners/suggestions, Graph and storage permissions/negative paths and result recording. All tenant cases are explicitly unexecuted. Field Customizer and Command Set coverage is explicitly export-only in this sample; no fake host mounting or simulated tenant pass is claimed. Permission approval, deployment and publication are separate authorized workflows.
- All old source links were relocated to the library workspace; obsolete line anchors were removed. Both browser storage exports point at the actual `useSPFxStorage.ts` implementation. Package README links repository docs through HTTPS, since those docs are not in the npm tarball.
- Provider/property docs now describe shallow snapshots and the requirement for host rendering, top-level removal/ref replacement, lack of nested in-place observation, explicit lifecycle cleanup, and external shared resources outside runtime-instance isolation. No universal memory or tenant guarantee is claimed.
- Storage docs now use the actual `SPFxStorageHook<T>` object result and object destructuring; old tuple signatures/examples were wrong. Defaults seed missing/unreadable current keys; default-only changes preserve current values; remove/external deletion use current defaults. Best-effort storage is distinguished from durable guaranteed persistence.
- OneDrive documents preserved lazy default seed, current missing-response ownership for automatic creation, latest read/write local-state ownership, independent pending channels and remote-write ordering limits. TenantProperty and TenantKV document their current error/fallback/loading/rejection behavior and identity/unmount guards, based on implementation and provider-storage/async-hooks reports.
- Tenant helpers/hooks/services consistently retain the legacy JSON/string format and disclose numeric/boolean/null-like string coercion, date strings, possible bigint precision loss and lack of runtime validation. No stored-data envelope or migration is promised.
- Existing references had additional pre-existing API drift. Environment, diagnostics, themes/container, user/photo/hub/list, permissions and PnP sections were reconciled from actual source declarations. Removed invented fields/types/methods include environment production/Test labels, Teams isInTeams/isTab fields, timeAsync/getMarks, photo url, permission canAdd helpers and search totalRows. Each of the 40 public hook sections and all four provider sections remains documented; helper and service inventories remain present.
- New fix descriptions were checked in current source: PnP expiry uses milliseconds with zero retaining five-minute legacy fallback, keyFactory identity refreshes SPFI, photo/cross-site target change guards, performance same-name concurrent timers, SPFx Teams wrapper/context replacement and precheck token-provider replacement. Placeholder recreation is included in the tenant lifecycle procedure. No guarantee was inferred beyond the implemented paths.

## Validation

| Check | Observed result |
|-------|-----------------|
| Existing red link evidence, `/private/tmp/spfx-doc-links-red.txt` | Public verifier failed for obsolete local source paths before alignment |
| `npm run verify:public-docs` on final documentation state | Exit 0: API inventory and relative link targets pass |
| `npm run verify:examples` on final documentation state | Exit 0: 40 hooks and 4 providers covered in sample registry |
| `git diff --check -- README.md packages/spfx-react-toolkit/README.md docs packages/spfx-react-toolkit/src/index.ts` | Exit 0, no whitespace errors |
| Revised self-contained TSX snippets | 16 snippets from both README files and rewritten environment/performance/theming/user-site/permissions/PnP references typechecked with local TypeScript, JSX React, ES2020/ESNext, node module resolution, strictNullChecks and workspace package; exit 0. Unique temp directory under `/private/tmp` was removed in finally. |
| Source/API documentation count | 40 `## useSPFx...` sections retained across ten hook category documents |

The public documentation verifier checks local link file existence and inventory; it does not validate every Markdown fragment or compile all illustrative legacy snippets. The targeted snippet compilation above is separate evidence and does not claim every remaining illustrative snippet was compiled. Root API/package/runtime/build checks belong to final integration verification; authenticated SharePoint/Graph checks remain unexecuted.

## Independent review correction round 1

The independent F7 review by `monorepo_phase_review` returned **FAIL with three P2 findings**. This round corrects those specific examples and requests re-review; no independent PASS is claimed here.

1. **Cross-Site Query captured the old base URL.** `setBaseUrl()` updates state for the next render; invoking the request in the same callback used the old `baseUrl`. The example now captures `targetUrl` and uses it directly for the GET, so the first click targets OtherSite. The example also handles its returned rejection and renders the hook error.
2. **Unavailable ServiceScope APIs.** Removed EventAggregator/SPPermission from the built-in services list and replaced the EventAggregator subscription example with `PageContext.serviceKey`, which is public in the installed SPFx 1.21.1 SDK. The custom-service example now includes the missing implementation class and complete React import, keeping the list and both snippets coherent.
3. **Application Customizer example repeated placeholder lifecycle defect B16.** It now subscribes/unsubscribes `placeholderProvider.changedEvent`, retries Top-placeholder creation, captures the placeholder owned by each disposal callback, clears only that current placeholder, renders a complete HeaderComponent, and safely unmounts/disposes its captured current placeholder on extension disposal. Old disposal callbacks cannot unmount the replacement container. This follows the current sample entry point's lifecycle pattern; tenant execution remains unperformed.

Validation on the corrected state:

- **4/4 complete revised TSX snippets typechecked**, using local TypeScript 5.3.3, installed SPFx 1.21.1 and the workspace package. Compiler options: noEmit, JSX React, target ES2020, module ESNext, Node resolution, skipLibCheck, strictNullChecks. Fresh isolated temporary directory was removed in finally. Evidence: [snippet typecheck](./evidence/documentation-review1-typecheck.txt).
- `npm run verify:public-docs`: **exit 0**, inventory and local file targets pass.
- `npm run verify:examples`: **exit 0**, 40 hooks and 4 providers covered.
- `git diff --check` for the three corrected docs and report: **exit 0**, no output. Evidence: [verification commands](./evidence/documentation-review1-verification.txt).

Only the three assigned public documents, this report and the two evidence text files were changed during correction round 1. No source/runtime/scripts/package changes, deployment, publication or new delegation occurred.
