# Maintain the documentation

Edit the source that owns a fact, then reconstruct and verify its public documentation. Source JSDoc owns API behavior and examples; editorial Markdown owns integration decisions, procedures and evidence.

[Documentation home](../INDEX.md) · [Development index](INDEX.md)

## Markdown formatting

Write each prose paragraph on a single source line and let the editor or renderer wrap it visually. Keep each list item's prose on one line, with separate lines for nested items or blocks. Preserve the line breaks required by code blocks, tables, headings and intentional Markdown hard breaks.

## Content ownership

| Content | Maintained source | Automatic output |
| --- | --- | --- |
| Product introduction and minimal example | Root `README.md` | Package `README.md` with repository links rebased; packaged LICENSE stays local |
| API purpose, behavior, requirements, selection guidance and complete examples | Owning public TypeScript JSDoc | `docs/api/**/*.md` and semantic API catalog |
| Home orientation | `docs/_sources/INDEX.md` | `docs/INDEX.md` with discovered navigation |
| Installation, first integration and provider choice | `docs/_sources/getting-started/**/*.md` | Matching `docs/getting-started/**/*.md` routes |
| Task-oriented choice, integration, migration and measurement reasoning | `docs/_sources/guides/**/*.md` | Matching `docs/guides/**/*.md` routes |
| Contributor setup, verification and authoring procedures | `docs/_sources/development/**/*.md` | Matching `docs/development/**/*.md` routes |
| Essential authenticated host checklist | `docs/_sources/development/sharepoint-validation/INDEX.md` | `docs/development/sharepoint-validation/INDEX.md` |
| Executable scenario modules and lazy catalog in the private sample | Owning public JSDoc examples with the executable contract | `components/generated/jsdoc-examples/*.tsx` and `catalogue.ts` |
| LLM discovery | Source metadata, discovered routes and package metadata | Root `llms.txt` |
| Dated reviews, results and evidence | Local `.docs/maintenance/` (plans in `.docs/superpowers/`) | None; never generator inputs |

Keep reusable test inputs in versioned `tests/fixtures/`, never local evidence. Public guides must remain useful without ignored files. Generated files carry an ownership marker and are not edited manually. A page containing human prose remains generated when that prose comes from an editorial source. The package README is generated from the root README, not maintained independently.

The CLI generates the complete public API/editorial snapshot from these owners. Update the source and regenerate the complete set; do not overwrite individual generated pages or remove unique content before its source ownership and link parity are verified.

## Source JSDoc contract

Document every public API and associated top-level public type with its purpose, observable behavior and a concrete valid example. Document every public input/member, required versus optional values, verified omission behavior and defaults, return values, callbacks/events, async ownership, errors, dependencies/providers and limitations. Identify inherited peer contracts and toolkit-specific differences. A source symbol list, a successful typecheck or a short replacement example does not preserve a unique integration recipe.

Distinguish **property absence** from **an explicitly supplied `undefined`** whenever runtime code checks own-property presence. An optional or `never` TypeScript spelling is not evidence of omission behavior. Verify discriminated branches independently: a shared member name or another branch's fallback does not establish its default. Add `@defaultValue` only for a confirmed applicable default and explain branch-specific behavior in the owning member/branch documentation. Preserve required legacy/peer `null` contracts with their reason; new absence state uses `undefined`.

Use standard TSDoc tags for `@param`, `@typeParam`, `@returns`, `@throws`, `@defaultValue`, `@deprecated`, `@example` and `@see`. Explain whether callbacks are notifications or awaited work, who catches rejections, which operations change pending/error state, and what cleanup or cancellation actually occurs. Do not infer runtime guarantees from types alone. Parameter tags describe the actual originating signature, including overload-specific inputs. Callable aliases and re-export bindings can have their own source documentation; preserve those local differences.

Use **one real `@remarks` block per comment**, with meaningful internal headings. The pinned parser keeps one remarks block; appending another can discard earlier content. Source-file namespaces use the owning module's prose. Runtime/type name overlap gets a visible ` (type)` heading on the type, preserving the runtime anchor. Generic documentation distinguishes lexical source-declaration parameters from current instantiated parameters; do not invent missing current generics to satisfy a template.

The only project-specific tags are:

- `@useWhen`: concrete conditions for choosing a directly consumer-selected control, hook, provider, helper or factory. Associated props, compound parts and returned methods share that context unless they present an independent choice.
- `@avoidWhen`: a real limitation or unsuitable use, with its reason and an existing alternative when useful. Omit it when no specific contraindication exists; omission is not universal suitability. Keep hard incompatibilities in requirements/limits too.
- `@docPage`: optional local route relative to `docs/api/`, without extension, query, fragment, absolute prefix or traversal. For example, `@docPage controls/pagination`. The default is the owning module path relative to `src/` with `.ts`/`.tsx` replaced by `.md`. Re-export-only barrels do not create a second owner. Conflicting assignments fail.

These tags are defined in the library's [TSDoc configuration](../../../packages/spfx-react-toolkit/tsdoc.json). There are no capability IDs, alias lists, guide registries or global example IDs.

## API authoring template

This complete example shows the source shape for an input-only helper. Verify its thresholds against the implementation when changing the contract. Keep detailed parameter/default constraints on their actual owners.

````typescript
/**
 * Classifies an explicit container width without measuring the DOM.
 * @useWhen
 * Convert an already measured width to the toolkit container category.
 * @avoidWhen
 * Use the container measurement hook when the React component must observe a host element.
 * @param containerWidth - Required width in pixels; categories use exclusive upper thresholds.
 * @returns The corresponding container category.
 * @example Classify an explicit width
 * ```ts
 * import { getSPFxContainerSize } from '@apvee/spfx-react-toolkit';
 * const category = getSPFxContainerSize(720); // large
 * ```
 * @docPage helpers/container
 */
````

Examples in actual API source must be complete and meaningful, with legal public imports, explicit dependencies and valid hook placement: call hooks unconditionally at the top level of React components or custom hooks, never at module scope, conditionally or inside arbitrary callbacks. They are typechecked from source with real project peers and no hidden declarations. TypeScript-only shape examples may illustrate composition without claiming they are mounted demonstrations. Interaction recipes must return observable UI and connect their actions. A React or native DOM event does not observe a returned Promise: catch rejecting actions inside the handler even when the hook also publishes error state, and render the relevant result/failure. Gate requests on the documented readiness/loading state; readiness can change before invocation, so a disabled button does not replace the catch. Keep initialization errors separate from request errors, and preserve fallback-returning operations rather than implying every failure rejects. Await application saves where the workflow depends on completion. Do not use undeclared application globals, `any`, assertions or private imports to make an example pass. Preserve unique late-placeholder, cleanup, CRUD, draft or migration sequences instead of replacing them with smaller examples.

Every public capability still requires the [AGENTS.md](../../../AGENTS.md) JSDoc, public Markdown and working test-web-part scenario deliverables. Generated reference satisfies the API Markdown portion; a new task explanation belongs in an editorial source. Update the sample registry/coverage where the capability changes, and report whether the sample actually ran. Documentation-only migration does not invent a new runtime scenario or tenant result.

## Editorial discovery and navigation

The resolver discovers `docs/_sources/INDEX.md` and recursively discovers Markdown under `docs/_sources/getting-started`, `docs/_sources/guides` and `docs/_sources/development`. Removing `/_sources` determines each public route. Start every source with an H1 and a useful introductory paragraph **before breadcrumbs**; those provide title and discovery summary. There is no required frontmatter, ID, manifest or per-guide registration. New focused guides must contain a real reader task rather than a symbol list.

Do not author `_sources/api`: API pages arise from the graph and local route tags. Development procedures and the essential authenticated host checklist are maintained in editorial sources and published as generated output. Every public Markdown page under `docs/` outside `_sources/` is generated; root README remains manually authored. The home adds ordinary navigation from discovered pages and links the API, development and SharePoint validation indices. Keep its authored content to orientation and a curated integration sequence; do not maintain another complete guide/capability/import catalog.

Use `{@link PublicName}` or `{@link package-specifier#PublicName | Label}` to reference a real public binding. Ambiguous short names fail. Relative Markdown links resolve from their **source location**, then rebase to the public destination; links between editorial sources resolve to their derived public routes. Public links must never expose `_sources` or `.docs`. Fenced, indented and inline literals remain unchanged.

## Editorial inclusion syntax

Use only the implemented finite inclusions outside literal code:

```markdown
<!-- toolkit:include symbol="PublicName" section="selection" -->
<!-- toolkit:include symbol="PublicName" section="example" -->
<!-- toolkit:include symbol="PublicName" section="example" title="Exact source title" -->
<!-- toolkit:include inventory="localization.resources" -->
```

`selection` renders normalized use/avoid guidance. `example` requires one unique local example or the exact title of a local source-owned example; multiple examples require unique local titles. It never silently chooses the first example. `title` is valid only for examples. Signatures, defaults, member catalogs and full API pages are not include sections.

The sole inventory variant reads [localization/resources.json](../../../packages/spfx-react-toolkit/localization/resources.json), schema version 1, and renders resource IDs/package-relative paths sorted by ID. It is mutually exclusive with `symbol`, `section` and `title`. Do not copy a manual resource table or introduce a second inventory language. Unknown/duplicate attributes, missing/ambiguous references, recursive inclusions and executable expressions fail. Ordinary HTML comments remain literal.

## Deterministic reconstruction and commands

The pure pipeline reads versioned package exports, tsconfig, TypeScript source/JSDoc, parser/URL constants, editorial sources and real package/localization metadata. It constructs one canonical graph and normalized contracts, validates all contracts and examples, renders API/catalog/README/editorial/LLM artifacts in memory, resolves provenance-aware links and checks every local public target/fragment before publication. It never rebuilds content from old generated pages, `lib`, ignored evidence or caches.

Identical source, configuration, lockfile and toolchain must produce the same paths and UTF-8/LF bytes, with one final newline for documents and authored blank lines preserved in extracted executable snippets. Generation uses no network, AI, clock, Git branch/HEAD/remotes or implicit environment values. Write versus check changes mutation policy, never content. Missing expected outputs are created; expected document routes may replace old unmarked content. Collisions, unsafe paths and source/link errors fail before writing, and unrelated unowned files are not adopted or removed. Only recognized generated outputs may be removed as orphans; unowned strangers remain protected. Retiring an obsolete unmarked page requires explicit reviewed migration deletion, never a broader orphan-deletion rule.

After authoring JSDoc, API or editorial sources, align both output families explicitly:

```bash
npm run generated:sync
# Review and version the source/output diff together.
npm run generated:check
# Run the build/tests appropriate to the change.
```

`generated:sync` invokes `node scripts/generated.cjs`; `generated:check` adds exactly `--check`. The unified command builds one canonical context, validates the complete source-example set once, renders docs and executable modules/catalog in memory, and preflights both destination policies before any writes. It accepts no other flags and ignores `TOOLKIT_DOCS_MODE` and `TOOLKIT_EXAMPLES_MODE`. Both families are checked for missing, changed and owned orphan outputs without changing bytes, mtimes or directories. The script locates the repository from its own file, independently of the caller's working directory.

The policies stay separate: documentation can replace an expected historical unmarked page, but executable output never replaces a manual file at an expected destination. Invalid canonical contracts, executable contracts, links or either policy preflight prevent changes to both families. Filesystem failures during application may leave earlier replacements applied; cross-family rollback is not promised. Source snippets retain authored blank lines after their marker; document formatting retains its one-final-newline rule.

The targeted commands remain available: `docs:generate`/`docs:check` call `scripts/documentation.cjs`, and `examples:generate`/`examples:check` call `scripts/executable-examples.cjs`. Only these targeted CLIs honor their respective `TOOLKIT_DOCS_MODE`/`TOOLKIT_EXAMPLES_MODE` values (`write` by default or `check`); unknown modes fail, and `--check` selects checking after validating the mode. The ordinary workflow uses the unified command instead of chaining these CLIs and repeating analysis.

The library build only checks localization inventory, compiles TypeScript and emits AMD resources. The app consumes the compiled library and versioned generated modules/catalog. A complete checkout builds without generation. Root/workspace build, serve, bundle, packaging and pack commands do not generate or check artifact freshness, including their npm lifecycle hooks. A missing or invalid generated module produces the normal app compiler error without hidden reconstruction. CI performs the explicit read-only generated check before build, without a global generation-mode override. Successful compilation and generated freshness are separate proofs. After library changes, keep sync → rebuild library → restart an active app serve; no generation at startup or new JSDoc watcher is installed.

Sync/check retain the cost of complete canonical analysis, while executable scenarios also add app compilation/bundling inputs even with lazy imports. Measure sync/check and library/app builds separately on the same environment; no fixed time or saving is promised. No command here authorizes publication, deployment, push or merge.

For scoped pipeline verification:

```bash
node --test tests/documentation-*.test.cjs tests/public-style-docs.test.cjs tests/form-docs.test.cjs tests/data-grid-docs.test.cjs tests/demo-groups.test.cjs
git diff --check
```

Both unified modes validate the **full unfiltered canonical model**, every source example and the final virtual public link set before comparing or writing. `verify:public-docs` uses the same graph and ArtifactSet without writing, retaining physical source-JSDoc, member-coverage and supported-import checks. Missing-member/import/JSDoc negative tests remain regression checks. Follow the [verification scope](verification.md#verification-scope) for additional gates. Run `npm run verify:docs-reproducibility` to prove cold reconstruction. The verifier creates two ordinary temporary input mirrors with the maintained development sources but without Git metadata, local evidence, compiled output, generated documents or the generated executable directory, and executes their copied engine scripts. Only installed package realpaths admitted by lockfile v3 are shared; workspace links and unlocked nested packages are rejected. It compares the complete path set, SHA-256 per file and a length-framed aggregate digest separately for documentation and executable families and for their union, covering virtual artifacts, actual written bytes and the current checkout outputs. Fresh Node processes vary both locales (`C`, `it_IT.UTF-8`), both timezones (`UTC`, `Europe/Rome`), discovery order, caller directory and input mtimes. The deleted-output phase removes the entire public development directory, verifies that its maintained sources remain byte-identical, and reconstructs every expected development path and byte. The same rebuild removes injected owned documentation and executable orphans and preserves manual strangers with their bytes and mtimes. Source mutation changes public JSDoc and the maintained development index, then converts and edits a copied executable source example, requiring the owning document and executable family to change. Source-escape probes remain separate.

Generation children preload the versioned offline guard, which rejects fetch, Node network APIs and child/worker launches. Deliberate negative probes establish that the guard is active. This checks engine dependencies; it is not a security sandbox for hostile code. Linux CI additionally requires `sudo` and `unshare`, captures the exact `setup-node` executable in `DOCS_NODE_BIN`, and fails when network-namespace capability is unavailable. It runs both an actual socket probe requiring `ENETUNREACH` and the canonical CLI in the same isolated network namespace:

```bash
sudo -n unshare --net -- sh -eu -c '
  "$1" tests/fixtures/documentation/reproducibility/network-namespace-probe.cjs
  exec "$1" scripts/generated.cjs --check
' docs-offline "$DOCS_NODE_BIN"
```

For a future clean checkout, run `npm ci --no-audit --no-fund`, `npm run generated:check`, `npm run build`, `npm run verify:docs-reproducibility` and the Linux namespace commands above. Installation may use the network; generation does not require it. Do not clean/reset an active user checkout. Report executed platforms separately: configuring Linux CI does not establish a Linux pass, and local macOS proof does not exercise Linux namespaces. These checks establish generated artifact byte identity, not universal SPFx bundle/tarball identity or authenticated SharePoint behavior.

## Moves and current links

Preserve public routes and useful natural headings where practical. When content moves, repair actual current incoming links in source JSDoc, editorial sources and root README to the meaningful new heading or canonical symbol. Rebase links from the original source. Do not maintain synthetic compatibility IDs, stale historical menus or a second route map.

Remove an old output only when every unique fact, complete recipe and evidence paragraph has real versioned ownership and final links pass. Local parity ledgers are review evidence, never generator inputs. Package/import/saved-data compatibility is separate from documentation anchors. Link/type checks do not execute examples, verify remote URLs, establish accessibility or prove authenticated SharePoint behavior; report those limits explicitly.
