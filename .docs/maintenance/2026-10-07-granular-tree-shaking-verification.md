# Granular tree-shaking final verification — 2026-10-07

**Status: implementation and fresh local development gates PASS; integrated review and scoped follow-up complete.** All final results below were produced after product/source stabilization and inspected from their completed reports. Authenticated SharePoint and real GitHub CI execution remain NOT EXECUTED. Release/version assessment remains required before publication.

## Scope and preservation

Work uses the selected `apvee/granular-tree-shaking` checkout. The library remains TypeScript ESNext/ES2020, with package version 2.1.0, `main: lib/index.js`, `types: lib/index.d.ts` and shared peer ownership. Exact legacy `/lib` aliases and canonical implementations/context identities are preserved; the clean styles facade re-exports those implementations. Feature-owning PnP services retain registration effects only when consumed. The standalone list leaf now includes its own webs registration.

Baseline source audit and measurements are retained in [the initial audit](2026-10-07-granular-tree-shaking-audit.md) and [initial production report](evidence/tree-shaking/initial/bundle-report.json). Optimized final candidate results are in [the completed installed-package evidence](evidence/tree-shaking/final/package/consumer-result.json). The initial audit's `passed: null` remains diagnostic; it is not converted into a release PASS.

The [independent final preservation check](evidence/tree-shaking/final/root-preservation-package-check.json) confirms unchanged name/version/main/types, dependencies/peers/devDependencies/engines, root lockfile, historical API/package baselines, frozen packed inventory and reviewed informational tree-shaking baseline. All **599 historical packed paths** survive. The **603-file, 243,856-byte** tarball adds only `lib/styles/index.js`, `index.js.map`, `index.d.ts` and `index.d.ts.map`, with **237 preserved documented declarations**. No source/app/test files or clean aliases for internals were added. The exact export map contains 763 keys; declaration compatibility covers 82 historical modules with the reviewed single webs-augmentation exception.

No commit, push, merge, publication, tenant permission change, deployment or dependency upgrade is authorized by this record.

## Fresh verification

The [final root command ledger](evidence/tree-shaking/final/root-commands.json), [repository log](evidence/tree-shaking/final/verify.log), [package log](evidence/tree-shaking/final/verify-package.log) and completed reports establish:

| Required check | Final status | Duration |
| --- | --- | ---: |
| `npm run build:library` | PASS, exit 0 | 1.420 s |
| `npm run build:app` after library | PASS, exit 0 | 10.016 s |
| `npm run verify` | PASS, exit 0; **663/663 tests**, both workspace typecheck/lint, docs/API/examples/runtime/import graph | 20.530 s |
| `SPFX_PACKAGE_EVIDENCE=.docs/maintenance/evidence/tree-shaking/final/package npm run verify:package` | PASS, outer exit 0; eight recorded command exits 0, no inner Gulp failure markers | 115.259 s |
| Final installed production bundle contract | PASS, complete 36 variants / 20 comparisons | 32.060 s, within package total |
| Final production runtime | PASS, complete 18 fresh probes | 29.949 s, within package total |
| Final negative mutations | PASS, all nine detected | 14.310 s, within package total |
| Packed manifest/export/peer/declaration preservation | PASS, historical contracts preserved | recorded above |
| Integrated review plus scoped final follow-up | No independently discovered Critical/Important; both final Minor findings ADDRESSED; no new scoped issues | reviewed snapshots |

The eight [package commands](evidence/tree-shaking/final/package/commands.json) are library build, isolated install, two dependency-tree checks, ship bundle, consumer TypeScript, strict public/deep declaration probes and solution packaging. Lint is covered by the fresh root `verify` in both workspaces; this record does not invent an additional consumer lint command. The actual consumer has been removed by normal default cleanup; [consumer-path.txt](evidence/tree-shaking/final/package/consumer-path.txt) remains an evidence record, not a reusable directory.

Both concrete integrated Minor findings are fixed: the source detector now distinguishes deferred object methods/accessors from invoked literal defaults/conditional returned factories; diagnostics preserve the installed consumer and clean only its recorded temporary directory. The regression tests failed for all three exact examples before the fix, then passed 16/16. A fresh source audit records 150 modules / 1,118 qualified exports, zero enforced source-rule issues and zero stale/missing/CommonJS outputs. Evidence: `final/task-7-detector-red.log`, `task-7-detector-green.log`, `task-7-import-graph.json`, `task-7-source-audit.log`.

The detector remains a syntactic guard, with no callable-alias/property-dispatch interpreter. Real emitted bundle/runtime gates provide independent protection. Development diagnostic snippets parse, and their guarded cleanup was executed against an exact scratch consumer. `git diff --check` passed after the two fixes.

The earlier [package-attempt-1 result](evidence/tree-shaking/final/package-attempt-1/consumer-result.json) and `final/verify-package-attempt-1.log` remain preserved as failed-attempt evidence. Its 36 bundle variants, 18 runtime probes and nine mutations passed, but sample ship/typecheck/solution failed because controller overlap copied the temporary `maxWidth.full` sample edit while Task 6's writer was live. The stable sample uses `rootWidth.full`. This is a preserved controller sequencing failure, not a current candidate failure; the complete fresh rerun above passed.

## Final production measurements

The [final production bundle report](evidence/tree-shaking/final/package/bundle/bundle-report.json) records the nine families below. Each cell is **raw / gzip9 / Brotli11 bytes**. For these 27 variants initial equals unique union, async is zero, and each root/domain versus leaf delta is **0 raw / 0 gzip / 0 Brotli**.

| Family | Root | Domain | Leaf |
| --- | ---: | ---: | ---: |
| Pure helper | 503 / 324 / 275 | 503 / 324 / 275 | 503 / 324 / 275 |
| Stable callback | 805 / 473 / 389 | 805 / 473 / 389 | 805 / 473 / 389 |
| Provider | 4,944 / 1,819 / 1,609 | 4,944 / 1,819 / 1,609 | 4,944 / 1,819 / 1,609 |
| Width descriptor only | 619 / 380 / 318 | 619 / 380 / 318 | 619 / 380 / 318 |
| PnP context | 27,586 / 9,529 / 8,570 | 27,586 / 9,529 / 8,570 | 27,586 / 9,529 / 8,570 |
| PnP list | 36,236 / 12,081 / 10,838 | 36,236 / 12,081 / 10,838 | 36,236 / 12,081 / 10,838 |
| PnP search | 22,134 / 7,552 / 6,724 | 22,134 / 7,552 / 6,724 | 22,134 / 7,552 / 6,724 |
| PnP batch | 21,544 / 7,426 / 6,684 | 21,544 / 7,426 / 6,684 | 21,544 / 7,426 / 6,684 |
| Minimum useSx | 34,220 / 11,747 / 10,552 | 34,220 / 11,747 / 10,552 | 34,220 / 11,747 / 10,552 |

All **18 blocking alias comparisons pass at zero gzip gap**, with no forbidden emitted-module retention. Required registrations/engine modules remain present. The width-only variants retain only one `width: 100%` descriptor call, without renderer/provider/typography/PnP retention. Minimum useSx retains width plus body1's fontFamily/fontSize/fontWeight/lineHeight, without other descriptor recipes or PnP/provider retention. Narrow callback and useSx probes equal their family totals.

Root gzip changes from the frozen fresh pre-change audit are:

| Family | Initial root gzip | Final root gzip | Change |
| --- | ---: | ---: | ---: |
| Pure helper | 13,321 | 324 | −12,997 |
| Stable callback | 13,459 | 473 | −12,986 |
| Provider | 14,732 | 1,819 | −12,913 |
| Width descriptor only | 13,376 | 380 | −12,996 |
| PnP context | 14,980 | 9,529 | −5,451 |
| PnP list | 14,906 | 12,081 | −2,825 |
| PnP search | 13,867 | 7,552 | −6,315 |
| PnP batch | 13,330 | 7,426 | −5,904 |
| Minimum useSx | 24,087 | 11,747 | −12,340 |

These changes compare equivalent named fixtures in the initial/final reports; they are informational absolute improvements, distinct from permanent alias-gap acceptance.

| Control | Initial raw/gzip/Brotli | Async raw/gzip/Brotli | Unique union raw/gzip/Brotli |
| --- | ---: | ---: | ---: |
| Empty / type-only, each | 412 / 267 / 227 | 0 / 0 / 0 | 412 / 267 / 227 |
| Direct Griffel | 28,184 / 10,015 / 9,090 | 0 / 0 / 0 | 28,184 / 10,015 / 9,090 |
| Already Fluent + useSx | 71,707 / 23,403 / 19,850 | 0 / 0 / 0 | 71,707 / 23,403 / 19,850 |
| Already Fluent direct | 66,421 / 22,007 / 18,689 | 0 / 0 / 0 | 66,421 / 22,007 / 18,689 |
| Dynamic styles catalog | 64,080 / 15,356 / 13,352 | 0 / 0 / 0 | 64,080 / 15,356 / 13,352 |
| Lazy root catalog | 3,437 / 1,724 / 1,482 | 261,180 / 66,256 / 55,084 | 264,617 / 67,980 / 56,566 |

The two engine comparisons are informational (`passed: null`): minimum useSx versus direct Griffel **+1,732 gzip bytes** (+6,036 raw / +1,462 Brotli); already-Fluent useSx versus already-Fluent direct **+1,396 gzip bytes** (+5,286 raw / +1,161 Brotli). They are not failed blocking comparisons. The [independent final asset check](evidence/tree-shaking/final/root-bundle-evidence-check.json) recompressed/hash-checked **37 emitted assets**, checked deduplicated unions and confirmed all nine equal triplets.

Budgets remain 1,023 gzip bytes for each root/domain versus leaf comparison, with the pure helper root budget 512 bytes. All forbidden-module groups must retain zero emitted members, independently of byte deltas. The canonical contract requires all 36 variants/comparisons. Audit subsets cannot pass the primary gate or write a baseline. Full/dynamic catalogs retain broader code by design; the lazy catalog warning is advisory, not a promise of property-level pruning.

Final provenance: Node **22.20.0**, npm **11.19.1**, TypeScript **5.3.3**, Webpack **5.95.0**, Terser **5.44.0**, React/ReactDOM **17.0.1**, SPFx **1.21.1**, PnP **4.17.0**, Griffel core/react **1.19.2 / 1.5.30**. Zlib is `1.3.1-470d3a2`, gzip level 9 / Brotli quality 11. React/ReactDOM and SPFx are host externals; PnP, Fluent and Griffel are bundled from the installed consumer only. Production enables minimization/usedExports/module concatenation/deterministic IDs and disables cache.

- Final tarball SHA-256: `6b19512f8a67af0f49abc37b6060fbb8b8dda8101da27a81973ed44130fc7299`.
- Root lock SHA-256: `56b44d939cac13a5b31cf9d81474bbc3c6b714da9b4438f4c7ffb7c509ed9dbf`.
- Final consumer lock SHA-256: `ff66d1054c4a64343cff962abc5848534a8eadb6a02527ef62d4c61578235744`.

The bundle/runtime reports retain installed origins, toolkit manifest/loader/contract hashes, per-template/fixture/config hashes and per-asset hashes. Only emitted chunk/nested-module membership establishes retained code; parsed orphans are excluded. Asset union deduplicates shared/initial/lazy assets. Original-baseline reproducibility is separately recorded in [initial/reproducibility.json](evidence/tree-shaking/initial/reproducibility.json), with zero cold-replay/relocation drift. The final evidence is a fresh complete run plus independent asset verification; no additional final cold replay is claimed.

## Runtime, mutations and warnings

The [final runtime report](evidence/tree-shaking/final/package/runtime/runtime-report.json) and [independent runtime/mutation check](evidence/tree-shaking/final/root-runtime-mutations-check.json) confirm all **18 probes** passed with exits 0 and no failure markers. Each scenario runs for root/domain/leaf in a fresh process, with registrations/client bundled in one realm and HTTP-send interception only:

| Scenario | Observed boundary |
| --- | --- |
| Context | Standalone and targeted web calls, headers, parsed results and resolved web URL |
| List | Query/CRUD, title/GUID/server-relative selectors, escaped names, batch creation and partial failures |
| Search | Search/suggest operations and parsed results, actual search prototype augmentation |
| Generic batch | Explicit initial batching absence, subsequent real registered batch operation |
| Providers | Canonical mixed-import context identity, two isolated stores, hook/host updates, unmount cleanup |
| Styles | Real Griffel/Fluent context identity, equivalent class/CSS rules, rerender deduplication, LTR/RTL and independent renderers |

The [nine final mutations](evidence/tree-shaking/final/package/mutations/mutation-report.json) all reject for their intended contract/runtime cause: unwanted PnP; unwanted width descriptor; missing new export; missing historical export; workspace-origin escape; byte inflation (helper root +17,503 gzip exceeds 512); lost list registration; lost batch registration; lost search prototype augmentation. The search mutation preserves valid query-builder imports while removing the actual bundled augmentation. Unrelated syntax/module-resolution failures cannot count as successful mutation detection.

The [resolver evidence](evidence/tree-shaking/final/package/entrypoint-resolution.json) verifies actual TypeScript **node/bundler/node16** modes, canonical declaration identity, clean/legacy resolution and generic inference; [shared package records](evidence/tree-shaking/final/package/shared-packages.json) verify consumer-owned Fluent/Griffel identity. Strict TypeScript/public/deep probes, independent SPFx ship bundle and solution packaging pass. Root lint covers both workspaces. Outer Gulp success cannot erase inner failure markers.

Warnings are retained in package commands/logs: existing transitive deprecations; npm install-script policy reports four entries (fsevents 1.2.13 twice, es5-ext 0.10.64, fsevents 2.3.3); baseline-browser-mapping and caniuse-lite metadata age advisories are explicitly routed and recorded by the existing guarded verification preload. No install-script approval, dependency upgrade or warning hiding was performed. Webpack records the lazy catalog's approximately 255 KiB async-asset advisory. Existing SDK root parse-stage browser failure is preserved with original browser failure evidence; the successful browser harness externalizes unused SPFx host modules using production conventions and asserts none are emitted. It does not replace toolkit hooks or mock the sample.

## Actual browser and host boundary

Task 6's actual sample browser evidence is `evidence/tree-shaking/task-6/browser/results.json` and `browser/final-report.md`, inspected for this record. **Chrome 155.0.8059.39: PASS, 65 observations and no failures**, zero console events/page errors/failed requests. Desktop 1440×1000 and mobile 390×844 exercised callbacks, width override/removal, equivalent classes/computed styles, FluentProvider LTR/RTL and focus/ArrowRight horizontal scrolling. Original overflow used a **390px viewport**, page scrollWidth 409px and a 360px preview override; the immutable integrated review's 360px viewport description is corrected here without editing that review. The fix contains previews in named keyboard-reachable horizontal scroll regions; library descriptors retain their intended values. Fixed mobile page scrollWidth is 375px against viewport 390px while the overridden preview retains computed width 360px. Initial failure evidence remains in `browser/red-browser-initial/`.

The sample browser fixture uses actual compiled root/domain/legacy imports with the actual ImportsPanel, independently of authenticated SharePoint. Chrome evidence does not establish a cross-browser matrix. **Authenticated SharePoint Imports and Styles validation: NOT EXECUTED.** No authentication, effective permissions, tenant deployment or SPFx-host lifecycle PASS is inferred from local results.

## Limits and development handoff

Verification applies to the pinned Node 22.20.0 / TypeScript 5.3.3 / React 17.0.1 / SPFx 1.21.1 / PnPjs 4.17.0 baseline and tested resolver modes. Native Node ESM/CommonJS package loading, the complete advertised peer-version range, cross-browser support and arbitrary dynamic namespace pruning are not verified. Existing renderer cache lifetime and dynamic CSS growth limitations remain documented.

The CI workflow invokes the single installed consumer gate and retains reports with `if: always()`, but **real GitHub CI: NOT EXECUTED** in this session. Local execution proves the recorded local gates, not a remote runner result.

**Ready for development review/handoff.** The completed fresh gates close the readiness condition in the immutable integrated and scoped reviews. The intentionally removed incidental root PnP registrations require a version/release assessment before publication, including consumers that formerly relied on unrelated root imports to augment their own PnP clients. Such consumers now own their explicit feature imports. Current package version 2.1.0 was preserved during implementation; this report does not invent a semver decision or authorize publication.

## Recorded rulings

All five rulings from the SDD progress ledger are preserved below; no additional ruling is inferred from routine fixes.

| Ruling | What and why | Cost if wrong / safeguard |
| --- | --- | --- |
| Preflight repository locations and immutable review | Repository paths override skill scratch defaults: SDD artifacts stay under `.docs/superpowers/sdd`, with no worktree or automatic commits. Immutable snapshots include untracked sources because no commit preserves evidence. | Additional local scratch disk; preserve this workspace/evidence. |
| Task 1 descriptor parser reuse | Extract the historical observation parser into the shared bundle helper, preserving the legacy `value` field/exact return format and normalizing `staticValue` only at the new report boundary. Avoid importing the old main's eager filesystem/stderr effects. | Legacy diagnostic format regression; focused format tests. |
| Task 2 declaration registration exception | Permit exactly one bare `@pnp/sp/webs` declaration import in the list service, with a module-specific guard and completion requirement. TypeScript emits registration augmentation although factory signatures remain unchanged. | A broad guard could hide API/registration regressions; duplicate/wrong-path/wrong-module negative tests and independent review. Historical baselines remain frozen. |
| Task 4 bare generic-batch client | Use `@pnp/sp/fi` plus default behaviors instead of broad `@pnp/sp`, whose RequestDigest path incidentally registers batching and masks the negative mutation. | Incomplete legitimate client setup; require positive operations and explicit pre-import batching absence in the same bundled PnP realm. |
| Task 4 interim dependency-tree reuse | For interim proof, extract a fresh current tarball into a copied real Task 3 dependency tree and verify hashes/origin/shared peers/SPFx. Avoid duplicate installation while dependency roles are unchanged; final Tasks 5/7 still require fresh normal package gates. | Installation regression might appear later; label interim evidence accurately and run the mandatory fresh normal gate, now completed. |

Command durations above are observed execution times, not model/token cost. No model/token cost totals were supplied, and none are invented.

## User confirmation and post-completion cleanup

On 2026-10-07 the user reported that the web part works correctly. This is a user-confirmed general functional result; environment, browser and individual host-checklist steps were not supplied. The earlier local automated/browser results remain agent-observed; detailed authenticated tenant checklist rows have not been individually recorded.

The user subsequently authorized deleting intermediate snapshots and duplicate proofs. A single compressed final snapshot now preserves all changed/new Git-visible source, fixture, configuration and documentation files, excluding generated evidence. Original Git HEAD and the original hash inventory remain available. Reports, decisions, initial/final evidence, failed-attempt summaries/logs and browser evidence remain preserved. Bulky raw webpack stats from intermediate Task4/Task5 and failed final-attempt runs were removed; their result/module/provenance reports remain. Final raw stats and all reusable tests/fixtures remain. The original per-task source states are no longer recoverable from local snapshots. Cleanup inventory: [cleanup.json](evidence/tree-shaking/cleanup.json).
