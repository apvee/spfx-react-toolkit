# Site collection key-value store verification

Date: 2026-10-06. Branch: dev. Base: cbd5ab3523492e6c9461f01b6a1d36093f3a8941. Implementation plan approved by user; commit and push to dev subsequently authorized on 2026-10-06. Final independent review PASS; no actionable findings.

## Result and scope

Added createSPFxSiteKeyValueStoreService and useSPFxSiteKeyValueStore for one hidden SiteKeyValueStore in the current collection root (pageContext.site.absoluteUrl), shared by subsites. Item CRUD/serialization is shared with tenant through an internal core; site-only schema/setup/paging logic is isolated. Public docs, Site sample and isolated consumer checks updated.

Tenant public signatures and all 82 historical declaration modules remain compatible. Tenant hook, catalog helpers/service, serialization helpers and legacy internal modules remain unchanged. Existing user AGENTS.md edit is preserved (SHA-256 checked against execution-start snapshot). Manifests, dependencies, peers, versions, lockfile and package allowlists are unchanged. The implementation/validation phase made no worktree, branch switch, commit, push, merge, publication, deployment or grant changes. Commit/push are handled afterward under the separate user authorization; the pre-existing AGENTS.md edit is excluded.

## Fresh integrated evidence

| Check | Result |
|---|---|
| npm run build:library | PASS, exit 0 |
| npm run build:app | PASS, exit 0 |
| npm run verify | PASS, exit 0; 203 tests, typecheck, lint, examples, runtime, docs and API |
| API compatibility | PASS; all 82 historical modules; only exact site hook/service barrel additions permitted and required |
| Demo coverage | PASS; 41 hooks and 4 providers |
| npm run verify:package | PASS, exit 0; 347 tarball files, 190096 bytes; isolated install, no workspace link, TypeScript public/deep contracts, shared peers, ship bundle and solution |
| git diff --check | PASS |
| Tenant protected files / baseline / manifests / lockfile | No diff |

The first package attempts hit default npm-cache write EPERM and restricted DNS ENOTFOUND. The restricted install was interrupted before package assertions. The final attempt used writable /private/tmp cache, fetch_retries=0 and automatically approved network escalation; no automatic approval rejection. All output/install files stayed in temporary directories, and no repository dependency changes resulted.

Baseline before implementation passed both builds and full verify (90 tests, 82 declaration modules). Existing browser metadata/dependency deprecation warnings and legacy expected error-path/MockTimers output remain; new workspace lint has no warnings. No toolchain upgrades performed.

## Tests and review corrections

Public factory/core characterization proves the original tenant REST traces, mutex/cache scope, URL formatting, single-page behavior, serialization/coercion, warning-only uniqueness setup, errors and method shape. Site tests cover collection isolation, read-only list/schema probes, only list-level 404 absence, strict setup/unique keys, grant masks/envelopes, complete validated pagination, race recovery, verified deletion absence, blank Note values and server collation.

React tests cover independent overlapping counters, newest errors, root/client identity changes including in-place PageContext mutation, stale grants/callbacks and unmount. Review regressions fixed original list-create error preservation through later schema failures and loading visibility when parent-provided operations start in child layout effects. Demo tests use the real hook/Fluent/shared UI and an external service boundary for failed reads/writes, empty values, retry, roots and unmount.

Task 1 and Task 2 reviews passed. Task 3 and Task 4 each passed scoped rereview after the documented fixes. Task 5 scoped compliance/quality and whole-change final review PASS. Reviewer read the complete product diff and fresh verification logs, reported no actionable findings, and did not rerun suites. Its declined checks are authenticated host behavior, an expanded SPFx version matrix, legacy tenant migration and external publication/admin work; those remain outside this local assignment, with limits stated here.

Local working records: ../superpowers/sdd/2026-10-06-site-key-value-store/ (ignored): progress.md, task reports/reviews, final-local-verification.log, final-package-verification.log and whole-change-review.diff. Implementation plan: ../superpowers/plans/2026-10-06-site-key-value-store.md; design: ../superpowers/specs/2026-10-06-site-key-value-store-design.md.

## Real-host limits

Authenticated SharePoint integration is NOT EXECUTED. Local controlled transports and package builds do not establish real root/subsite ACLs, tenant endpoints, runtime UI or remote ordering. Eleven separate host scenarios and disposable-key cleanup/evidence instructions are in ../../docs/SHAREPOINT-VALIDATION.md#site-collection-key-value-store. They are not local failures or deployment authorization.
