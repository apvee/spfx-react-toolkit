# useSx implementation results — measured cost

Final post-correction measurements, 2026-10-07. This report records local generated
JavaScript, real renderer CSS/cache observations and fresh repository/browser gates.
The final scoped-composition correction is independently reviewed. Authenticated SharePoint validation
is **NOT EXECUTED**.

## Reproducible setup

Run `node scripts/measure-sx-bundle.cjs` from the repository root. The frozen
baseline is `artifacts/sx-baseline/library.tgz`, SHA256
`d87c017581e27d58d2525be9133041953341dc85f9ea61601d68ba92303fa7d5`.
It is extracted directly and never rebuilt from candidate HEAD. Candidate
tarballs, JS, Webpack stats, command logs, hashes, module attribution and CSS
snapshots are archived in `.docs/maintenance/evidence/use-sx/task-7-final-measure/`.
Final candidate SHA256:
`b1b5af2ea1db82e7bdcbe1f8e420dafa29bb85e39b8a9ec69c59c2041b5efb02`,
matching the separately verified package consumer tarball.
Preoptimization evidence and its tarball are in `task-6-preoptimization-stable/`.
The initial preoptimization tarball remains in the original measurement output
and was reused after annotations, rather than recreating a supposed old source.

Locked actual versions: Node 22.20.0, npm 11.19.1, TypeScript 5.3.3, Webpack 5.95.0,
React 17.0.1, Griffel core 1.19.2/react 1.5.30, Fluent theme 9.2.0 and shared
contexts 9.25.2. The evidence includes lockfile, host config and frozen host
manifest hashes. Every paired fixture uses identical compiler options and
externals: React/ReactDOM and SPFx are supplied by the host; Fluent/Griffel are
bundled. A stable temporary resource path avoids changing deterministic module
IDs merely by changing the evidence label. Module concatenation is disabled in
the isolated probes to make module attribution inspectable; the separate actual
SPFx ship bundle uses its existing production configuration.

Sizes below are bytes: raw generated JS, gzip at level 9, Brotli at quality 11.
Each reachable emitted chunk is counted once. Maps and license sidecars are
excluded. Compression sizes are summed per asset, matching separately served
chunks. They are not the compression of one concatenated archive. Module source
sizes in stats are before minification, not per-module compressed contributions.

## JavaScript cost

| Fixture | Archived before PURE raw / gzip / Brotli | Final post-correction raw / gzip / Brotli |
| --- | ---: | ---: |
| Existing helper, frozen baseline | 46,006 / 15,191 / 13,479 | 46,006 / 15,191 / 13,479 |
| Same existing helper, candidate; no useSx | 107,451 / 22,276 / 19,305 | 80,403 / 19,677 / 17,085 |
| Isolated minimum UI, direct Griffel | 64,538 / 15,754 / 13,693 | 64,538 / 15,754 / 13,693 |
| Equivalent minimum UI, deep hook/helper wrapper imports | 98,256 / 20,510 / 17,555 | 71,772 / 17,907 / 15,516 |
| Equivalent minimum UI, package-root wrapper imports | 143,478 / 35,144 / 30,295 | 116,994 / 32,459 / 28,281 |
| Existing FluentProvider + toolkit app, direct UI | 186,792 / 48,710 / 40,683 | 159,744 / 45,817 / 38,594 |
| Equivalent existing FluentProvider + toolkit app, wrapper UI | 188,950 / 49,551 / 41,489 | 162,466 / 46,742 / 39,434 |
| Full interactive StylesPanel/catalog | 458,728 / 125,863 / 103,392 | 458,952 / 125,917 / 103,329 |

The archived before-PURE column precedes the final native scoped-binding correction.
That correction adds 224 raw bytes to used-wrapper fixtures; before/after annotation
evidence remains separately archived. The final package-root minimum delta over
direct Griffel is **52,456 raw / 16,705 gzip / 14,588 Brotli bytes**.

The equivalent minimum UI applies width 100% and the four official body1 font
properties to the same div. Isolated deep imports add **7,234 raw / 2,153 gzip /
1,823 Brotli bytes** over direct Griffel. In the fixture that already has
FluentProvider, webLightTheme and an existing package-root toolkit helper, the
wrapper adds **2,722 / 925 / 840 bytes** over the equivalent direct UI.
These are measured fixture deltas, not universal costs or performance claims.

Package-root imports preserve legacy PnP registrations and toolkit entry-point
effects. The larger root-only minimum comparison therefore includes those
effects; it is not the isolated style engine cost. Stats retain PnP registration
modules for both historical/candidate root imports and none for the deep style
consumer. No package-wide `sideEffects:false` was added.

## What selective imports actually remove

Before annotations, the optimized unused and minimum consumers still contained
all **360** facade data-constructor call sites, including **140** recipe/spacing
calls. Forty-two facade modules now mark only their static data initializers,
and the nested static declaration constructors, with `/*#__PURE__*/`.
`createDeclaration`, `createRecipe` and static `createSpacing` copy/freeze data;
they do not register PnP, read DOM, use a renderer or insert CSS. Numeric `.px`
and `grid.columns` validation and all renderer/insertion/cache code remain
unannotated. No engine, Griffel `makeStyles`, or CSS insertion call was marked
pure. The annotations do not change runtime behavior when a member is used.

Actual optimized-JS AST observations, in addition to Webpack used-export data:

- Candidate no-useSx consumer: **zero** facade descriptor constructor calls.
- Both minimum wrapper imports: **one** body1 recipe, its four font declaration
  constructors, and **one** width declaration with value100%. No width.auto or
  other facade recipes survive those selective consumers.
- Full UI intentionally consumes the dynamic catalog, so it retains the catalog
  and its selectable members. Its size does not fall from the annotations; the final binding correction adds the small separately measured code delta.
- **All436 unique Fluent token variable references and all 17 upstream typography
  entries remain** in these candidate and direct-Fluent bundles. This is upstream
  catalogue retention, distinct from wrapper descriptor-member retention. The
  no-useSx root consumer still grows by **34,397 / 4,486 / 3,606 bytes** over the
  frozen baseline after the wrapper allocation fix. No property-level pruning
  promise or zero-cost unused import claim is supported.

The source and emitted-call inventories are recorded per bundle, with module
paths and used exports. These observations establish retained generated code,
not execution counts, heap bytes, browser parsing time or render latency.

## Actual CSS and cache growth

Real React 17/ReactDOM mounts/renders use a real Griffel DOM renderer and its
real stylesheets/insertionCache in JSDOM 15. There are no CSS mocks. The CSS
byte count is the UTF-8 sum of actual serialized rule strings joined by one
newline. JSDOM observations do not establish browser computed styles, native
scrollbar appearance or authenticated host behavior.

| Stage, one renderer | Rules | CSS bytes | Insertion entries | Binding factories | Assignment factories |
| --- | ---: | ---: | ---: | ---: | ---: |
| Before mount | 0 | 0 | 0 | 0 | 0 |
| Cold mount width 240/body1 | 150 | 17,819 | 150 | 5 | 5 |
| 100 identical renders | 150 | 17,819 | 150 | 5 | 5 |
| 100 renders alternating 240/360 | 151 | 17,868 | 151 | 5 | 6 |
| 1000 new distinct widths 1000..1999 | 1151 | 68,878 | 1151 | 5 | 1006 |
| After component unmount | 1151 | 68,878 | 1151 | 5 | 1006 |

The five property bindings materialize the engine's scoped binding rules;
therefore this cold mount inserts more than five rules. Repeated values reuse
classes/factories. The 1000 distinct values add **exactly 1000 persistent rules
and 1000 assignment factories**; they add 51,010 serialized CSS bytes in this
probe. Unmounting removes neither the renderer CSS nor its cache entries.
The WeakMap isolates renderer instances and permits unreachable renderers to
be collected; it does not evict values, remove CSS, or bound growth while a
renderer remains alive. Renderer garbage collection and heap consumption were
not measured.

## Actual SPFx ship assets and warning status

The selection starts from current archived/emitted manifests, adds their exact
path resources, and follows the current entry's Webpack `.u` chunk-name/hash
tables. Baseline `currentBuild` metadata corroborates the run. Repeated Webpack
builds preserve mtimes of identical lazy assets, so mtime-only selection is
insufficient. For example, the frozen entry explicitly calls `.e(877)` and
references its exact archived hash even though that shared lazy chunk's
`currentBuild` flag is false. It is included once because it is proven reachable;
other old unreferenced JS and all maps are excluded.

| Actual ship asset union | Assets | Raw | gzip | Brotli |
| --- | ---: | ---: | ---: | ---: |
| Frozen baseline | 16 | 563,032 | 169,661 | 143,653 |
| Candidate emitted assets | 18 | 702,451 | 201,446 | 169,664 |
| Candidate minus baseline | +2 | +139,419 | +31,785 | +26,011 |

This full-sample delta includes the actual new interactive UI, FluentProvider,
catalog, and chunk/module graph changes. It is not an isolated useSx cost.

The frozen baseline's wrapper returned 0 while Gulp reported exit 1 from the
existing browser-metadata age warning. It remains a **reported failed ship**.
A candidate parent-only preload retry likewise emitted assets after successful
tsc/lint/webpack but returned 0 while Gulp reported exit 1 and stderr failure.
That failure/status and warning log are preserved in
`task-6-ship-warning-failure/`; the measurement script returned 1, demonstrating
that it does not accept wrapper status alone. The final replay uses the
separately tested `scripts/gulp-build-metadata-preload.cjs` guard in the Gulp
parent and inherited worker environment. Only the two exact known age advisories
are routed visibly to stdout with a verification metadata label; every other
warning/error/raw stderr keeps its original behavior. It does not change locked
versions, compiler options, browser targets or runtime bundles. Final actual
tsc/lint/webpack/bundle completion: **process exit 0, zero failure markers,
two visible advisory lines, clean ship status**. The final full measurement
command exits 0. A further source-level retention-assertion replay explicitly
skips ship and does not replace this separate actual ship result.

## Verification boundary

- Fresh library build, strict bundle-fixture TypeScript check and candidate pack:
  exit 0, no failure markers.
- Eight production Webpack probes: emitted without errors or missing-export
  warnings; the advisory about locked baseline-browser-mapping data remains
  visible and is captured in the report.
- Descriptor/catalog/layout tests:37/37 pass. Library lint:exit 0.
- `git diff --check`:exit 0.
- API/package compatibility and final repository/browser gates are owned by
  their respective Task6/Task7 workstreams and must be reported separately.
- SharePoint/tenant validation:NOT EXECUTED.


## Initial Task 7 verification — archived before composition correction

The following archived checks were executed after the Task 6 metadata resolution correction but before the final review discovered the foreign scoped-class P2. They are retained as history; the refreshed post-correction evidence below supersedes them.

| Check | Fresh result | Evidence |
| --- | --- | --- |
| `npm run build:library` | Exit 0 | `evidence/use-sx/task-7-build-library.log` |
| `npm run build:app` | Exit 0; compiler/lint/webpack completed | `evidence/use-sx/task-7-build-app.log` |
| `npm run verify` | Exit 0; 573/573 tests, both workspace typecheck/lint, example/runtime/docs/API gates | `evidence/use-sx/task-7-verify.log` |
| `npm run verify:package` | Exit 0; all 8 isolated consumer commands exit 0 with no failure markers, actual ship/typecheck/probes/solution | `evidence/use-sx/task-7-package/commands.json` and `task-7-package.log` |
| `npm run bundle:ship` with guarded verification preload | Exit 0; actual Gulp compiler/lint/webpack completed, no failure markers, age advisories visible | `evidence/use-sx/task-7-bundle-ship.log` |
| Rebuilt core/catalog browser fixture | Chrome 152.0.7977.83: 586/586, no console/page errors | `evidence/use-sx/task-7-core-browser.json` |
| Rebuilt actual StylesPanel browser fixture | Chrome 152.0.7977.83: 473/473, no console/page errors | `evidence/use-sx/task-7-panel-browser.json` |
| `git diff --check` | Exit 0 | Root tool observation |

The isolated package contains 599 files, 236,221 packed bytes and 237 public JSDoc declarations. The root measurement replay `evidence/use-sx/task-6-root/measurement.json` matches all final comparison totals. A separate root recompression check verified every selected ship asset and the 16/18 unique-asset totals.

The prepared ship command is:

```sh
NODE_OPTIONS='--require /Users/fabiofranzini/GitHub/apvee/spfx-react-toolkit/scripts/gulp-build-metadata-preload.cjs' npm run bundle:ship
```

The corresponding package verifier installs the same guarded preload for its isolated Gulp commands. This is a verification-only decision: exact known locked browser-data age advisories are visibly classified on stdout; every other warning/error/raw stderr and explicit failure marker remains fatal. The risk is that the old browser metadata remains old, so the advisory cannot establish a current cross-browser compatibility matrix. No dependency, compiler target or runtime library logging was changed.

`task-7-integrity.json` confirms the original AGENTS.md, analysis, useStableCallback source/demo/tests remain byte-identical; all transitive lock entries are unchanged. Only the two approved workspace metadata entries differ. Branch remains `apvee/fluent-style-utilities`, HEAD `3135cd4735af969ee234b42c720d34ee6fc2ce20`, ordinary `.git` checkout. No commits, worktrees, pushes, merges, publication or deployment occurred.

Authenticated SharePoint-host validation remains **NOT EXECUTED**. Local real browser, isolated package and solution generation establish their stated boundaries only.

## Final composition correction and refreshed root evidence

The final whole-assignment review reproduced a native foreign Griffel hover width of 400px defeating a later toolkit hover width of 360px. The native property now materializes in the actual descriptor query/state scope, with the same shared variable priority chain, unconditional local resets and canonical scope in the renderer binding-cache key. This preserves atomic native property/scope identity and an earlier foreign base value outside inactive scopes. No public types, provider, hash/insertion ownership or catalog contract changed.

Regression evidence: runtime 7/21 before → 21/21 after; full behavioral suite 587/587. The independent scoped review approved the correction with no residual finding. Final refreshed measurement: `evidence/use-sx/task-7-final-measure/measurement.json`, actual ship status 0/no failure markers; the candidate tarball hash above is the post-correction artifact.

Fresh root `npm run verify` passes **587/587** and all typecheck/lint/examples/runtime/docs/API gates (`task-7-final-verify.log`). Fresh library build is recorded by the final measurement and isolated package commands; fresh app build is `task-7-final-build-app.log`. Fresh isolated `npm run verify:package` passes all eight commands, including actual ship, typecheck, root/deep probes and solution generation (`task-7-final-package/commands.json`). Package: **599 files / 236,426 bytes / 237 preserved public JSDoc declarations**.

The final root browser replays pass on **Chrome 155.0.8059.39**: **793/793** core/catalog/composition observations and **473/473** actual-panel observations, with zero console/page errors (`task-7-final-core-browser.json`, `task-7-final-panel-browser.json`). The browser executable updated between earlier Chrome152 runs and the final runs; no project toolchain dependency changed. Fresh guarded `npm run bundle:ship` exits0 with actual Gulp tsc/lint/webpack completion and zero failure markers (`task-7-final-bundle-ship.log`). Fresh `git diff --check` exits0.

The final independent whole-assignment review and one scoped fix rereview approve **local handoff with no open findings**. Earlier records and failures remain archived. All seven local phases are complete. Authenticated SharePoint remains **NOT EXECUTED**; no tenant-ready or wider-browser assertion is made.

The branch and initial source changes remain preserved in the same checkout, with no commit, push, merge, publication or deployment. The verification-only age-advisory classification above is the single tooling ruling; if an age advisory should be a strict build blocker, run without that preload and accept the separately recorded locked-metadata failure. Other diagnostics remain fatal.
