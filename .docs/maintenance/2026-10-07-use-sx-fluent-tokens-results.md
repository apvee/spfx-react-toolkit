# useSx — direct Fluent token references

2026-10-07; baseline PR #6 at `7be9285`, branch `apvee/fluent-style-utilities`.
This record covers the local uncommitted token refactor. Authenticated SharePoint
validation is **NOT EXECUTED**.

## Change and equivalence

Eight facade modules now import `tokens` from the existing
`@fluentui/react-theme` peer: foreground, background, presets, border-color,
border-width, border-radius, box-shadow and scrollbar. All **88 accesses / 38
unique tokens** produce exactly the prior CSS variable strings in the installed
theme 9.2.0 / tokens 1.0.0-alpha.22. Scrollbar uses
`${tokens.colorNeutralStrokeAccessible} transparent` and preserves its native
width, fallback and forced-colors behavior.

Spacing already used official tokens; typography keeps official
`typographyStyles`. Native CSS values, private `--apvee-sx-*` variables, exported
signatures, renderer/cache/scope behavior, all 119 affected PURE annotations and
useStableCallback are unchanged. Manifest and lockfile hashes match the baseline;
React 17.0.1, TypeScript 5.3.3 and SPFx 1.21.1 are preserved.

The new AST convention guard rejects handwritten theme references in production
literal/template values and verifies official direct accesses. It failed before
the source edit with 88 violations, then passed. Existing literal behavioral
oracles remain byte-identical. Public Markdown now explains source tokens versus
emitted CSS. Existing JSDoc and observable sample scenarios remain accurate.

## Measured JavaScript delta against the current PR

Run sequentially with the same stable paths and locked configuration:

```sh
SX_MEASURE_LABEL=tokens-before node scripts/measure-sx-bundle.cjs
SX_MEASURE_LABEL=tokens-after node scripts/measure-sx-bundle.cjs
```

Before/after tarballs, optimized JS, stats, command results, reachable SPFx assets
and CSS snapshots are archived in `evidence/use-sx/tokens-before/` and
`evidence/use-sx/tokens-after/`. The original frozen pre-useSx tarball remains
untouched. Paired values below are raw/gzip/Brotli bytes; compression is per asset
at gzip level 9 and Brotli quality 11. Maps/license sidecars are excluded.

| Probe | Before | After | Delta |
| --- | ---: | ---: | ---: |
| No useSx, current package root | 80403 / 19677 / 17085 | 82811 / 19924 / 17343 | +2408 / +247 / +258 |
| Minimum direct Griffel | 64538 / 15754 / 13693 | 64538 / 15754 / 13693 | 0 / 0 / 0 |
| Minimum wrapper, root | 116994 / 32459 / 28281 | 119402 / 32802 / 28536 | +2408 / +343 / +255 |
| Minimum wrapper, deep | 71772 / 17907 / 15516 | 74180 / 18227 / 15825 | +2408 / +320 / +309 |
| Already Fluent + toolkit, direct UI | 159744 / 45817 / 38594 | 162152 / 46160 / 38825 | +2408 / +343 / +231 |
| Already Fluent + toolkit, wrapper UI | 162466 / 46742 / 39434 | 164874 / 47085 / 39656 | +2408 / +343 / +222 |
| Full StylesPanel catalog | 458952 / 125917 / 103329 | 458605 / 125905 / 103315 | −347 / −12 / −14 |
| Actual reachable SPFx ship assets | 702451 / 201446 / 169664 | 702096 / 201430 / 169563 | −355 / −16 / −101 |

The selective probe increase is explained by **88 additional token property
reads** retained in optimized JS. The minifier removes the PURE descriptor calls
but preserves their property-reading arguments: an unused foreground module
still contains expressions such as `o.L.colorNeutralForeground1`. The AST
attribution in `tokens-after/token-access-attribution.json` observes 68 official
member reads before and 156 after in the no-useSx probe. No new descriptor
allocations survive there; both minimum wrappers still retain exactly six
constructor calls. All 436 upstream token variable names and 17 typography
entries remain, as before. Neither zero-cost unused imports nor property-level
token pruning is claimed. No global getter-purity or sideEffects override was
introduced to hide this cost. Full-catalog and actual ship output are slightly
smaller; this is a measured fixture result, not a universal saving.

CSS/cache snapshots are identical at every stage: cold width240/body1 has 150
rules / 17819 bytes; 100 identical renders add none; alternating width240/360
has 151 / 17868; 1000 distinct widths has 1151 / 68878, retained after unmount.
These use real Griffel with JSDOM CSSOM; browser proofs are recorded separately.

Before candidate SHA256: `b1b5af2ea1db82e7bdcbe1f8e420dafa29bb85e39b8a9ec69c59c2041b5efb02`.
After candidate SHA256: `24f034a5acb60a68f46c0023784b484941800e09b78d86a65fdad0535270c208`.

## Fresh local gates

Evidence is under `evidence/use-sx/tokens-validation/`; root tool-session exit
statuses are recorded with provenance in `command-status.json`.

| Gate | Actual result |
| --- | --- |
| Focused catalog/layout/descriptor tests | 39/39, exit 0; independent root rerun |
| Library then app build | Exit 0; actual tsc/lint/webpack/Gulp completed |
| `npm run verify` | Exit 0; 589/589 tests plus both workspace typecheck/lint and examples/runtime/docs/API |
| `npm run verify:package` | Exit 0; eight isolated commands status 0, empty failure-marker arrays; ship/typecheck/public+deep imports/solution |
| Tarball contract | 599 files / 236522 compressed tarball bytes / 237 preserved JSDoc declarations |
| Core/catalog browser | Chrome 155.0.8059.39; 793/793 observations; no console/page errors |
| Actual StylesPanel browser fixture | Same Chrome; 473/473; no console warnings/errors or page errors |
| Both actual measurement ship builds | Exit 0; zero failure markers; guarded metadata age advisories visible |
| Diff whitespace | `git diff --check`, exit 0 |

Browser fixtures were rebuilt after the library/app builds. The existing cached
Playwright was used because the Browser plugin was unavailable; no dependency
was installed in the repository. Core served on localhost4317 and panel on4320.
Pointer/keyboard actions, desktop/mobile, light/dark/high-contrast, forced-colors,
inheritance/preset overrides, composition/reset, RTL and container/viewport
behavior were exercised. Root inspected JSON assertions and desktop/mobile
panel screenshots. The panel uses actual sample code and compiled library with
local SPFx boundary substitutes. It establishes no tenant authentication or
SharePoint-host pass. Existing platform-dependent scrollbar appearance,
unsupported inverted high contrast and arbitrary-value CSS retention remain.

Task 1, Task 2 and the complete final measurement/source package received
independent clean compliance/quality reviews with no actionable findings. Root
also inspected the actual patch, reran focused tests, checked archived assertions
and measurement comparisons, and verified that the measured tarball matches the
isolated verified package. Verification was completed locally before the user's
subsequent commit/push request. No worktree, merge, publication or deployment is
part of this change.
