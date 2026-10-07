# RESULT

SOURCE AND RUNNERS STABLE. Restored Fluent scrollbar thumb hover and pressed/drag feedback while retaining 6px width/height, 3px rounded caps, transparent track/corner, native thin fallback, forced-colors system behavior and foreign Griffel order semantics.

# CHANGES

- scrollbar.ts: descriptor carries official colorNeutralStrokeAccessibleHover and colorNeutralStrokeAccessiblePressed tokens; JSDoc names exact base/hover/pressed behavior.
- descriptor.internal.ts: optional scalar scrollbarThumbHover/scrollbarThumbPressed fields; no new public exports/options or CSS injection API.
- bindings.internal.ts: existing supported vendor/forced-colors:none branch emits thumb:hover, thumb:active and more-specific thumb:hover:active. Combined selector gives pressed priority while hovered regardless of insertion order. No theme import or renderer/global/listener initialization.
- resolve.internal.ts: binding cache includes each new state color independently. Existing native binding/vendor variable composition design retained.
- tests/sx-catalog.test.cjs: real Griffel CSS state/color/pressed priority and independent metadata-cache regressions. Official theme access count extends from1 to3 for the approved added tokens.
- docs/api/helpers/styles.md, docs/SHAREPOINT-VALIDATION.md, StylesPanel.tsx: exact6px/3px feedback behavior and observable actual-thumb interaction guidance.
- scripts/check-sx-browser-fixture.cjs and scripts/check-styles-panel-browser.cjs: real pointer over thumb, pressed, drag, release, leave for both axes; drag scroll assertions; per-phase screenshot captures with expected themed RGB values, pointer coordinates and geometry. CSS selector support checked. Paint verification is explicitly marked as requiring screenshot/pixel review, since getComputedStyle(el,'::-webkit-scrollbar-thumb') reports the base declaration regardless of native widget state.

# EVIDENCE

- RED before production edits:2 new focused tests failed for missing hover selector/color and missing metadata-sensitive cached state colors.
- GREEN: node --test tests/sx-catalog.test.cjs tests/sx-runtime.test.cjs passed42/42; /private/tmp/scrollbar-hover-focused-tests.log. Existing6px/3px, foreign native composition, query/state and forced-colors tests retained.
- node --test tests/public-style-docs.test.cjs tests/demo-sx.test.cjs passed8/8; /private/tmp/scrollbar-hover-docs-demo-tests.log.
- npm run typecheck passed both workspaces.
- node --check both browser runners passed.
- git diff --check passed.
- Installed token verification: webLightTheme base#616161 hover#575757 pressed#4d4d4d; webDarkTheme#adadad/#bdbdbd/#b3b3b3; teamsHighContrastTheme#ffffff/#1aebff/#1aebff.
- Root separately confirmed CSS.supports for thumb:hover, thumb:active and thumb:hover:active in Edge154/Chrome155 and baseline paint reproduction with missing feedback. Root owns those external browser observations.

# ISSUES

Root owns final fresh library/app builds, actual Chrome/Edge pointer/paint validation, independent review and final npm verify/package gates. This agent did not rerun full npm test or package verification to avoid duplicate gates. Current runners capture real painted widget states but do not certify those colors solely from computed pseudo-element styles. Horizontal native paint may need separate root investigation in the current macOS environment; do not treat geometry/CSS as horizontal-paint proof. Authenticated SharePoint validation NOT EXECUTED. No new agents, branch/worktrees, commits or pushes.

# RUNNER DIAGNOSIS AND CORRECTION

Root's first full Generic browser runs failed4 dark-theme drag assertions per browser. This was preserved as failing evidence, not skipped or downgraded. The library was not changed in this correction.

A minimal source fixture reproduction isolates the cause: inserting a fullPage screenshot after light-theme interactions leaves the following dark native scrollbar with computed vendor6px styles but client320x100 versus offset320x100 (physical gutter0). Pointer hit-testing then targets catalog-scroll-content instead of the native scrollbar, and drag scrollTop stays0. Without that fullPage screenshot, light→dark→highContrast→dark gives client314x94 for the same320x100 box, and every vertical drag scrolls66/67px. Moving the overview screenshot after all interactions preserves all four passes. The browser is installed Chrome155 using real native widgets; diagnostic script /private/tmp/scrollbar-dark-runner-diag.cjs and /private/tmp/dark-diag-*-before/after.png.

Corrected only scripts/check-sx-browser-fixture.cjs: retain per-phase element thumb captures during real interactions; defer the normal full-page theme overviews to a capture-only pass after all actual thumb pointer tests. Full original overview paths remain. Dark drag assertions stay unchanged. Root's independent fresh actual StylesPanel dark probe on Edge/Chrome reported exact173idle→189hover→179pressed, supporting the isolated runner cause.

Fresh syntax/whitespace checks pass. Root owns the full corrected browser rerun and144-frame paint verification. Native full-page screenshot viewport mutation is a browser harness limitation: inspect physical gutter and hit-testing when CSS reports valid vendor styles but widget interactions unexpectedly fail. Computed CSS alone does not establish a painted or interactable native scrollbar.

# OVERSIZED CATALOG CAPTURE FOLLOW-UP

The first capture-order correction was incomplete: root's corrected full browser runs still had the same4 dark drag failures per browser because the normal-theme #catalog detail screenshot remained between themes. That catalog is taller than the900px viewport; Playwright also expands the viewport for oversized element screenshots. Individual100px thumb screenshots were not the offending captures.

The final runner correction defers both normal-theme full-page overview and the entire oversized #catalog detail capture to the final capture-only pass after actual thumb interactions in all3 themes. Forced-colors overview/detail captures already occur after all real thumb interaction tests. No oversized screenshot remains in the normal-theme pointer loop. All original capture paths, dark drag assertions, scoped/forced/foreign checks and production code remain unchanged. Fresh node --check scripts/check-sx-browser-fixture.cjs and git diff --check pass. Browser closure and pixel certification remain root's pending final rerun; the earlier repeated failure is not represented as a pass.
