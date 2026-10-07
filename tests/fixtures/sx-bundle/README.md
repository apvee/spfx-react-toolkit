# useSx cost probes

Run `node scripts/measure-sx-bundle.cjs` from the repository root. It builds the
library, checks these TypeScript consumers, packs a candidate without lifecycle
scripts, extracts the frozen baseline and candidate, and uses the installed
Webpack/TypeScript/React/Fluent/Griffel versions. It installs nothing.

- `unchanged`: identical existing helper import against both tarballs.
- `direct` and `wrapper-deep`: isolated equivalent width/full + body1 UI,
  comparing direct Griffel with the generic published hook/helper deep imports.
- `wrapper`: the same UI through the package root, including the root's existing
  PnP registration effects.
- `fluent-direct` and `fluent-wrapper`: identical FluentProvider + webLightTheme
  host that already consumes the toolkit, with equivalent minimum UI.
- `full`: the actual interactive sample StylesPanel and its catalog, consuming
  the candidate library through the package root.

React17/ReactDOM and SPFx modules are host externals consistently across every
probe. Fluent/Griffel are bundled consistently. These fixtures measure generated
code, not an authenticated SharePoint session. The full sample's ship build is
captured separately with the real SPFx host manifest/config and lazy chunks.

The script sums each reachable JS chunk once, with gzip9 and Brotli11, excluding
source maps and license sidecars. Stats include source sizes before minification,
module paths and used exports; emitted-JS AST observations establish actual
descriptor-call retention separately from upstream token/typography data.
Per-module source sizes are not a compressed-byte allocation.

Environment options:

| Variable | Behavior |
| --- | --- |
| `SX_MEASURE_LABEL` | Evidence label; letters, numbers, `_` and `-` only |
| `SX_MEASURE_OUTPUT` | Evidence base directory |
| `SX_MEASURE_CANDIDATE_TARBALL` | Reuse an archived candidate, e.g. for preoptimization comparison |
| `SX_MEASURE_SKIP_SHIP=1` | Intermediate code/CSS measurement only; explicitly records skipped ship |
| `SX_BUILD_METADATA_PRELOAD` | Verification-only guard path; installed in parent and worker environment |

Run sequentially: `temp/sx-bundle/current` is a stable resource path so evidence
labels do not change deterministic Webpack module IDs. Each evidence directory
archives JS, stats, candidate tarball and command logs. A failed command or Gulp
reported failure produces a nonzero script exit even when the child OS status is
zero. Emitted ship assets can still be measured with an explicit failure status.

The CSS probe uses real React17/ReactDOM and a real Griffel DOM renderer in JSDOM15,
without CSS mocks. It records actual rules, insertion-cache entries and factory
cache sizes before/cold/identical/repeated/1000-distinct/unmount steps. It does not
establish browser computed CSS, rendering performance, heap consumption, garbage
collection of discarded renderers, or tenant behavior.
