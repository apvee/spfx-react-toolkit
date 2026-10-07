# Task 3 local-browser observations

The independent expectations in `tests/fixtures/sx-browser/expected.json` were written before any cascade corrections. No cascade corrections were necessary. The existing Griffel assignment/reset selectors, fixed fallback chain and host comparator remain unchanged.

## Final environment

- Node 22.20.0; React/ReactDOM 17.0.1; TypeScript 5.3.3; Webpack 5.95.0.
- Griffel core 1.19.2/react 1.5.30; FluentProvider 9.22.8; shared-contexts 9.25.2; theme 9.2.0.
- Cached Playwright 1.64.0-alpha-1790635538000; installed Chrome 152.0.7977.83, headless. Native scrollbar-hiding default argument removed.
- Local fixture URL `http://127.0.0.1:4318`; initial desktop viewport 1210×900, plus 479/480/639/640/1024 widths during state/media checks and 390×844 mobile screenshot.
- Browser plugin unavailable. Cached Playwright library used in one process; no dependency installation or CLI daemon persistence assumption.
- SharePoint-host validation: **NOT EXECUTED**.

## Computed values and interactions

Final complete observations are in `task-3-browser-final.json`: **259 assertions, all passing**, no page errors and no console warnings/errors. The second fresh page reverses both descriptor order and mounted responsive-node order. Each order returns identical values.

The orchestrator's independent replay is preserved separately as `task-3-root-browser.json` and matching screenshots/DOM captures: 259 assertions, zero failures, empty console-message list, same Chrome version. It independently inspected the threshold-640 screenshot.

| Case | Computed observation |
| --- | --- |
| Main parent 639 → 640px | `display:flex`; `flexDirection:column` → `row`; two child rectangles prove vertical → horizontal placement |
| Width at 479/480/639/640/1023/1024px | 100/180/180/240/240/360px |
| Gap at 639/640/1024px | 4/20/32px |
| Explicit viewport with parent size changes | 420px throughout while actual window remains 1210px |
| Viewport 639 → 640px | width 120px → 420px; hovered viewport-state background red → green |
| Nested outer 1050px, inner 479px | outer child 360px; inner child 100px, using nearest named ancestor |
| Independent 639px/640px regions | column/row concurrently |
| Container element itself | own 1024px remains 1024px; its responsive width 111px never activates from itself |
| Equal medium container and viewport declarations | container width 260px wins over viewport width 250px |
| Parent hover and child's own declarations | parent width 600px; child width remains 75px, color `rgb(36, 36, 36)`; child medium private variable is locally invalid/empty |
| Omitted child color | inherits ordinary parent color `rgb(11, 22, 33)` |
| Independent sx results and strings, both orders | base 360px, hover 240px; independent container 280px and container+hover 320px survive resets |
| External Griffel composition | intercalated width 400px; later descriptor 360px; external last 400px; descriptor last 240px |
| Scope removal while pointer remains inside | width 333px → 111px after rerender; native `:hover` remains true |
| Base states | hover red, pointer active green; keyboard focus-visible blue wins during actual hover+Space active |
| Container states | focus-visible purple wins during all three simultaneous states |
| State responsive priority at 479/480/640/1024px | hover 110/140/200/260px; active 120/150/210/270px; focus-visible 130/160/220/280px |
| Native keyboard outline | 1px, `auto`, `rgb(0, 95, 204)`; present during simultaneous states, query states and selected enabled state |
| Enabled → disabled with pointer on button | native disabled attribute true, hover remains true, background becomes gray `rgb(100, 100, 100)` after enabled recipes are removed |
| Selected preset change with pointer on button | `aria-pressed:true`, orange `rgb(200, 100, 0)`; enabled hover restores red when selected is removed |
| Vertical `writing-mode:vertical-rl` | height/inline-size 639 → 640px gives column → row while physical width remains 160px |
| RTL logical padding | inline-start/right 13px, left 0px, direction rtl |
| RTL physical private padding | Griffel flips physical left 17px to physical right 17px |
| sx-owned probe/buttons | no inline `style` attributes |

The state width matrix sets viewport and container sizes to each matching threshold, distinguishing equal-threshold container priority from unconditional large viewport rules. The other responsive matrix holds the viewport fixed while resizing only the container. Private recipes exercise the generic engine without exporting the future flex/gap/background/preset catalog.

Ten final screenshots and matching DOM captures use the prefix `task-3-browser-final-`: forward/inverse threshold-640, simultaneous-states, selected-focus, RTL and mobile. Forward threshold and selected-focus screenshots were visually inspected, together with the earlier expanded simultaneous-state screenshot. The fixture is a test surface with intentional vertical container height and narrow width probes, not a product layout.

## Failures and diagnostic corrections

- `task-3-container-red.log`: 6/7 tests passed; missing public container facade assertion failed as intended. Adding the facade produced 7/7 in `task-3-container-green.log`.
- First fixture typecheck rejected `foreground.primary`, which belongs to the later catalog stage (`TS2339`). Replaced that fixture use with a private declaration retaining `var(--colorNeutralForeground1)`; no catalog expansion or type widening.
- Initial local server command `node scripts/serve-sx-browser-fixture.cjs` failed, exit 1: `Error: listen EPERM: operation not permitted 127.0.0.1:4317`. Retrying with `require_escalated` was automatically approved. Browser launches were also automatically approved. No automatic approval rejection occurred.
- `task-3-browser-initial.log/json`: 175 checks, four failing assertions in both orders. The token assertion compared literal `#242424` to equivalent computed `rgb(36, 36, 36)`. Its oracle now normalizes the token through a temporary ordinary DOM element. The scope-removal assertion used a centered pointer that fell outside the reduced width. It now positions the pointer at x+5 within both widths; it still asserts actual hover.
- `task-3-pointer-diagnostics.json` records the geometric cause: before hover width 222px at x=25, pointer x=136; hover width 333px; removed scope width 111px ends at x=136, with hover false after two animation frames. `task-3-fixture-diagnostics.json` is the earlier immediate diagnostic; its hover read happened before layout/hit-test settling and is not the final pointer observation.
- `task-3-browser-harness-corrected.log/json`: 175/175 passed after those harness fixes, with engine unchanged.
- `task-3-browser-expanded.log/json`: 235/235 passed after adding the larger combined-state matrix and focus checks.
- `task-3-browser-final.log/json`: 259/259 passed after adding actual flex-child geometry assertions. Final favicon endpoint returns HTTP 204; no console errors are filtered.
- Fixture build reports the existing stale `baseline-browser-mapping` warning. Dependencies were not upgraded or warnings suppressed.

## Verification commands and exits

| Command | Exit / evidence |
| --- | --- |
| `node --test tests/sx-runtime.test.cjs` before facade | 1; container red log |
| Same command after facade | 0; 7/7, container green log |
| `node scripts/build-sx-browser-fixture.cjs` final | 0; `task-3-final-build-fixture.log` |
| `npm test` | 0; 483/483, `task-3-final-root-tests.log` |
| `npm run typecheck` | 0; both workspaces, `task-3-final-typecheck.log` |
| `npm run lint` | 0; both workspaces, `task-3-final-lint.log` |
| `npm run build:library` | 0; `task-3-final-build-library.log` |
| `git diff --check` | 0; `task-3-diff-check.log` |
| Inline baseline integrity probe | 0; `task-3-integrity.json` |

Exact final browser command, run through an automatically approved escalated exec:

```bash
PLAYWRIGHT_MODULE_PATH=/Users/fabiofranzini/.npm/_npx/31e32ef8478fbf80/node_modules/playwright SX_BROWSER_EXECUTABLE='/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' SX_BROWSER_URL=http://127.0.0.1:4318 SX_BROWSER_RUN_LABEL=task-3-browser-final node scripts/check-sx-browser-fixture.cjs > .docs/maintenance/evidence/use-sx/task-3-browser-final.log 2>&1
task3_exit=$?
tail -40 .docs/maintenance/evidence/use-sx/task-3-browser-final.log
exit $task3_exit
```

Result: exit 0. Prior invocations used labels `task-3-browser-initial`, `task-3-browser-harness-corrected` and `task-3-browser-expanded`, with exits 1, 0 and 0 respectively. Initial/harness-corrected used the default port 4317. Final/expanded used port 4318.

Active server command: `SX_BROWSER_PORT=4318 node scripts/serve-sx-browser-fixture.cjs`, exec session **87450**. It is deliberately left running for the orchestrator's independent replay. The obsolete port-4317 server was verified as PID 6564, `node scripts/serve-sx-browser-fixture.cjs`, and stopped with `kill -TERM 6564` through approved exec; that owned server session 75487 ended with exit 143. No unrelated process was terminated.

Integrity probe confirms that among all files present in Task 3's baseline, only the planned styles index, runtime test and Development protocol changed. All existing engine files, comparator, package manifests/lockfile and previous unrelated modifications remain byte-for-byte intact. Frozen tarball SHA256 remains `d87c017581e27d58d2525be9133041953341dc85f9ea61601d68ba92303fa7d5`.

Full `npm run verify`, `npm run verify:package`, public catalog docs and executable SPFx sample remain later feature gates. No authenticated SharePoint-host, other-browser or tenant-permission result is claimed.

## Review correction: documentless scope-removal regression

The orchestrator/reviewer found an uncaptured JSDOM CSSOM parse diagnostic in the new scope-removal runtime test. `node --test tests/sx-descriptors.test.cjs tests/sx-runtime.test.cjs` initially returned exit 0, 12/12, but printed `There was a problem inserting the following rule` for `@container apvee-sx (min-inline-size: 640px)` with `--apvee-sx-width-container-640:222px`, followed by `Error: Unexpected } (line 1, char 103)`. Full output is preserved in `task-3-review-noise-red.log`.

Added the missing initial container-assignment assertion before changing the harness; the same scoped command failed with exit 1, 11/12, `the initial container assignment must be observable`. Full red output: `task-3-review-removal-red.log`.

Changed only `tests/sx-runtime.test.cjs`: its removal test now mounts real React with the real documentless Griffel renderer and filters `renderer.insertionCache` by mounted class tokens. It proves container 222px and hover 333px assignments exist before rerender, both are absent afterward, and base 111px remains. This checks class removal and emitted rules without involving JSDOM's unsupported container parser. No console errors are suppressed; no production code or browser expectation/evidence changed.

The scoped command passed exit 0, 12/12 in `task-3-review-removal-green.log`, then passed fresh again in `task-3-review-removal-final.log`. Final output was explicitly checked for insertion/parser/error/warning diagnostics and contains none. `git diff --check` also passes. The earlier root-suite/build/static results remain evidence for the pre-review checkpoint; this final correction is test-only and was verified with the requested bounded descriptor/runtime suite.
