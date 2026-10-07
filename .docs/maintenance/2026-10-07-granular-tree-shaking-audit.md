# Granular tree-shaking source audit and fresh baseline — 2026-10-07

Task 1 audit on `apvee/granular-tree-shaking`. No library, application, dependency, package manifest, export map or purity annotation changed. This is an audit baseline (`passed: null`), not a release gate pass.

## Complete inventory

[`initial/import-graph.json`](evidence/tree-shaking/initial/import-graph.json) contains all 149 source modules, every static import/re-export edge, entrypoint-qualified canonical leaves and source/build checks. [`initial/export-inventory.json`](evidence/tree-shaking/initial/export-inventory.json) joins all 1,118 qualified entrypoint exports to their canonical source, runtime peer dependencies, direct importers and manual effect classification. Qualified namespace members are counted individually; repeated exports through different entrypoints are retained.

| Entrypoint | Runtime leaves | Type leaves |
| --- | ---: | ---: |
| `index.ts` | 320 | 121 |
| `core/index.ts` | 4 | 10 |
| `hooks/index.ts` | 47 | 58 |
| `services/index.ts` | 11 | 33 |
| `helpers/index.ts` | 258 | 20 |
| `helpers/styles/index.ts` | 229 | 7 |

The collector uses the TypeScript 5.3.3 checker to follow named aliases/imported exported bindings and namespace qualification. Whole-program ESNext emit distinguishes erased type imports/re-exports, including unmarked aliases and `export type *`. The existing `collectPublicSurface` is reused as a cross-check where supported; its inability to follow the imported `useSPFxContext` binding is recorded for root/hooks rather than hiding or duplicating the declaration compatibility checker.

`utils/index.ts` explicitly exports no public utility symbols; all utility source files are still audited and their historical packed paths are frozen.

The fresh library build has 149 JavaScript modules, 149 declarations and their maps (596 files). No missing output, stale JavaScript/declaration without source, or CommonJS emission was found. `tsconfig.json` remains ESNext/ES2020/node resolution. The frozen tarball contains 599 files, including three package metadata/document files, and excludes source/app/config/tests.

## Manual effect review and next-task ownership

| Owner | Module-evaluation behavior | Required action |
| --- | --- | --- |
| `core/context.internal.tsx`, `core/state.internal.tsx` | Allocate React contexts; development displayName assignments mutate those local contexts. Runtime stores are created by provider instances. | Preserve shared context identity and provider isolation. No external registration identified. |
| Style leaf modules and `descriptor.internal.ts` | Descriptor/recipe factories allocate and freeze local data; no renderer/CSS/DOM work at import. | Pure call-site work in later tasks must preserve descriptor semantics. |
| `styles/types.ts`, `cache.internal.ts`, `bindings.internal.ts` | Allocate a brand Symbol, per-renderer WeakMap and lookup tables. | Retain shared identity and renderer ownership. Do not infer automatic purity from static analysis. |
| `spfx-api-permission-precheck.helpers.ts` | Allocate local audience/status lookup Sets. Requests and diagnostics are deferred to exported functions. | Preserve behavior; no external module-evaluation registration found. |
| `spfx-pnp-context.service.ts` | Bare imports of `@pnp/sp/webs` and `@pnp/sp/batching`. | Task 2 moves registrations to actual feature owners. Context currently imports batching despite not using a batch implementation. |
| `spfx-pnp-list.service.ts` | Bare imports of lists/items/batching; standalone source does not import webs. | Task 2 must provide its missing webs registration rather than inheriting it incidentally from root/context. Not fixed in Task 1. |
| `spfx-pnp-search.service.ts` | Bare import of search. | Preserve search registration when the feature is consumed. |
| `spfx-pnp.service.ts` | Bare import of batching. | Preserve batch registration when the feature is consumed. |

No source runtime cycles or internal imports from directory `index` barrels were found. Unknown effects must remain review-required; static observations are not a purity certificate. The only identified external module-evaluation registrations are the four known PnP owners above. No new external-effect checkpoint is needed before the approved next task.

## Fresh tarball consumer measurement

The existing strict package verifier installed the current tarball into an isolated external SPFx consumer and passed TypeScript, root/deep import compatibility, shared Fluent/Griffel package identity, ship bundling and solution packaging. The packed inventory is frozen in `tests/fixtures/package-entrypoints.json`; it must not be regenerated from post-change output.

Node/npm provenance is recorded as v22.20.0/npm 11.19.1. npm metadata for the saved baseline/replay/relocation was captured after the runs and independently confirmed by the original package verifier notice; `initial/npm-provenance.json` records that timing. Future runs capture npm before building. Asset facts are unchanged.

Tarball SHA-256: `24f034a5acb60a68f46c0023784b484941800e09b78d86a65fdad0535270c208`.

The production collector uses the consumer as Webpack context, generates exact-slot entries in `.tree-shaking/entries`, resolves bundled peers only from its node_modules, and validates toolkit origin. React/ReactDOM and SPFx are host externals; PnP, Fluent and Griffel are bundled. Production uses minimization, usedExports, module concatenation, deterministic IDs and `cache: false`. Complete emitted chunk/nested-module IDs and reasons are retained losslessly in `*-stats.json.gz`; parsed orphan inventories are excluded. The readable report stores emitted resources/reason counts, unique JS assets and initial/async/union raw/gzip9/Brotli11 sizes/hashes.

All nine family templates retain observable exports/UI; provider variants consume the same child context hook. Width root/domain use `width.full`, legacy leaf uses `full`, and incoherent width slots are rejected. Clean domain aliases and the stable/useSx probes are explicitly unavailable until Task 3. Missing legacy imports, errors, truncated stats, missing reachable chunks, wrong package origin or missing enforce variants are fatal even during audit.

| Family | Root gzip bytes | Legacy leaf gzip bytes |
| --- | ---: | ---: |
| helper | 13321 | 324 |
| stable-callback | 13459 | 473 |
| provider | 14732 | 1819 |
| sx-width-only | 13376 | 380 |
| pnp-context | 14980 | 9529 |
| pnp-list | 14906 | 12076 |
| pnp-search | 13867 | 7552 |
| pnp-batch | 13330 | 7426 |
| sx-minimum | 24087 | 11747 |

25 variants are measured; 11 planned clean aliases are unavailable. The 28 recorded contract observations describe the pre-change overhead/registration retention. Comparison limits of zero are diagnostic placeholders, not approved final budgets; Task 5 sets reviewed limits from the fresh baseline.

Empty and type-only controls emit identical assets (267 gzip bytes). Minimal styles retain exactly body1 plus width/fontFamily/fontSize/fontWeight/lineHeight (six constructor call sites); width-only retains width=100% (one site). Upstream token object properties are not toolkit constructor sites and do not imply property-level token pruning. Griffel-direct is 10,015 gzip bytes versus minimum leaf 11,747 (+1,732). Already-Fluent direct is 22,007 versus toolkit leaf 23,403 (+1,396).

The genuine dynamic catalog retains eight sites. The lazy root namespace is an intentional broad control: its feature is consumed in an async chunk, with 1,724 initial and 66,256 async gzip bytes (67,980 unique union); it retains 384 sites, rather than claiming minimal descriptor pruning. Its 255 KiB async asset triggers Webpack's size warning, retained in all reports. The runtime's outdated baseline-browser-mapping advisory is retained in command logs; no dependency update was performed.

## Reproducibility and limitations

[`initial/reproducibility.json`](evidence/tree-shaking/initial/reproducibility.json) records zero drift across the complete cold replay and a copied-installed-consumer relocation for helper root/leaf, minimum styles root/leaf and lazy. Asset names/hashes/raw/gzip/brotli, generated fixture/config hashes, emitted resource lists and retention observations match. Tarball/toolchain/lock/loader/compression provenance is identical; only consumer/toolkit absolute locations are excluded from the path comparison. No new installation occurred for relocation. Initial/replay/relocation costs were 25.781/32.168/8.613 seconds respectively.

The relocation copy was removed after saving evidence; the original kept package consumer remains available for subsequent tasks/independent inspection (see `initial/cleanup.json`). No authenticated SharePoint-host or browser execution is claimed; Task 1 measures production bundles and checks types/package/local tests. Behavioral feature probes are owned by Task 4.

The first npm package attempt failed because the sandbox denied default home npm cache writes (EPERM, misleading root-owned-files text). A task-local `/tmp` cache and approved network execution completed the existing verifier. No home permissions were changed. An early measurement's truncated nested stats report is archived as invalid; regression coverage now rejects truncation. Provisional reports are archived separately and are not the baseline.
