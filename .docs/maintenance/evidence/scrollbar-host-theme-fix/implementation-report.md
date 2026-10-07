# RESULT

SOURCE STABLE. SPFx host-theme conversion now preserves the shim's base accessible stroke and every other converted token while mapping accessible-stroke hover to host neutralPrimary and pressed to host neutralDark. Undefined SPFx fallback and all shared Teams theme identities remain unchanged. No scrollbar geometry, renderer, public signature/export, dependency or generic theme mapping changes.

# CHANGES

- packages/spfx-react-toolkit/src/helpers/spfx-theme.helpers.ts: bounded2-token override after the real createV9Theme conversion. Optional host palette fields are read safely; missing, empty or whitespace-only state colors preserve their upstream converted fallback. Input theme/palette is not mutated. Source JSDoc documents behavior, limitations and usage.
- tests/spfx-theme-helpers.test.cjs (new root test): real Fluent UI8 createTheme + real createV9Theme establishes collapsed upstream state colors, default host expected literals, custom inverted palette values, source immutability, equality of every other converted token, missing/blank state fallback and undefined/Teams theme identities.
- tests/demo-sx.test.cjs: fixture accepts a real host theme at the unavailable SDK boundary; new test mounts the real hook, helper, StylesPanel and FluentProvider in host mode and verifies scoped CSS variables '#605e5c' base/'#323130' hover/'#201f1e' pressed. Existing undefined-host checks remain.
- docs/api/helpers/INDEX.md: dedicated existing theme helper section documents exact3-token palette mapping, missing-value fallback, custom/inverted/equal-color limits, peers and FluentProvider usage.
- docs/api/helpers/styles.md: SharePoint host mapping note connects the existing scrollbar state selectors to the restored scoped theme values.

# EVIDENCE

RED before production change:
- Real helper regressions failed because hover '#605e5c' differed from desired host '#323130', custom inverted hover '#a1a2a3' differed from '#b1b2b3', and valid partial state colors were ignored.
- New actual panel/provider host-mode regression failed for the same '#605e5c' versus '#323130' CSS variable discrepancy. An initial test harness getComputedStyle naming error was fixed before confirming the intended failing assertion.

GREEN:
- node --test tests/spfx-theme-helpers.test.cjs tests/demo-sx.test.cjs tests/public-style-docs.test.cjs passed13/13; /private/tmp/scrollbar-host-theme-tests.log.
- npm run typecheck passed both workspaces after safely reading optional IReadonlyTheme.palette (initial compiler error was resolved without casting away optionality).
- git diff --check passed.

Installed upstream v9ThemeShim maps base/hover/pressed accessible stroke to palette.neutralSecondary. Default real v8 palette supplies neutralSecondary '#605e5c', neutralPrimary '#323130', neutralDark '#201f1e'. Root's read-only real authenticated Edge workbench inspection had observed all3 CSS vars '#605e5c'; root owns the actual-host provenance and final reread.

# ISSUES

Root owns fresh build/library + app gates, restart of the known running serve process, actual user-page CSS/paint confirmation, independent review, full verify/package gates and PR work. This agent did not restart serve, claim authenticated tenant permission coverage, push, commit, delegate or change branches/worktrees. Actual browser fixture SDK still uses undefined fallback by default, but the new DOM sample/hook/provider host-mode test consumes a real host palette and the root will verify the actual user page. Custom palettes can intentionally use equal colors; no universal contrast/distinctness guarantee is claimed.
