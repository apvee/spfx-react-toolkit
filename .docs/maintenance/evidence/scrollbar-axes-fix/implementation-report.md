# RESULT

FINAL INTEROPERABILITY FIX: SOURCE AND RUNNERS STABLE: scrollbar.fluent uses the user's final approved 6px width and height with 3px rounded thumb caps. This supersedes the earlier 8px prototype/implementation. Vendor styling is scoped through Griffel. A per-property/scope private vendor variable selects auto inside the supported branch; the native property stays at its original Griffel atomic scope, preserving both argument/merge orders with foreign native declarations. Local initial resets block inherited vendor assignments, including inactive scopes. Standard fallback remains thin; forced-colors retains native thin sizing/system colors. Public exports/imports/peers, provider identity and explicit overflow/container dimensions/gutter/overscroll contracts remain unchanged.

# CHANGES

- scrollbar.ts: final 6px private metadata, separate neutral thumb token, public JSDoc/example.
- descriptor.internal.ts: optional binding fields scrollbarSize ('6px') and scrollbarThumb (single CSS color).
- bindings.internal.ts: vendor pseudo width/height6px via metadata and 3px thumb radius; supports selector(::-webkit-scrollbar) + forced-colors:none branch assigns auto through a private variable consumed by the ordinary native binding. Existing query/state wrapping and local inherited-variable resets retained.
- tests/sx-runtime.test.cjs: new real Griffel foreign width:none/color:red regressions cover segmented sx and mergeClasses in both orders for base, hover and container-hover.
- resolve.internal.ts: complete binding cache key includes both metadata fields.
- tests/sx-catalog.test.cjs: real Griffel emitted CSS regressions assert final6px both axes, 3px caps, forced-colors isolation, independent cache orders and responsive/state selector scope.
- docs/api/helpers/styles.md and docs/SHAREPOINT-VALIDATION.md: final6px vendor/fallback contract and explicit bidirectional example/checklist.
- StylesPanel.tsx: dedicated scroll preview has explicit bidirectional overflow and large content; final6px explanatory text. Existing every-catalog scenario remains bidirectional.
- Browser runners check final6px axes, neutral thumb, horizontal scrolling, inactive scoped appearance, recipe removal/restoration and forced-colors native axes. sx fixture now includes bidirectional scrollboxes, removable recipe and inactive child beneath styled parent.

# EVIDENCE

Final interoperability correction:
- RED: all3 new real Griffel order/scope regressions failed before production change because the vendor branch retained native auto declarations after later foreign overrides.
- GREEN: node --test tests/sx-catalog.test.cjs tests/sx-runtime.test.cjs passed40/40; /private/tmp/scrollbar-fix-interop-tests.log.
- npm run typecheck --workspace @apvee/spfx-react-toolkit passed after correction.
- Existing browser fixture/runner extended for independent native hiding/recoloring, both sx/merge orders, real hover and inactive hover, plus explicit3px rounded thumb check.
- Docs explain native non-auto overrides may bypass6px vendor sizing in Chromium.

Prior final6px sizing evidence:
- RED: node --test --test-name-pattern='fluent scrollbar emits' tests/sx-catalog.test.cjs failed with 'The vertical scrollbar uses the approved 6px width' before changing production from8px.
- GREEN: node --test tests/sx-catalog.test.cjs passed16/16; log /private/tmp/scrollbar-fix-6px-catalog.log.
- node --check both changed browser runners: passed.
- git diff --check: passed.
- Targeted search: no remaining8px scrollbar contract in affected source/docs/runners/tests/sample.

Prior verification history for the superseded8px version (not final6px gate evidence):
- Initial RED missing vendor axis rules:3 regressions failed,13 existing passed.
- Focused sx catalog/normalization/runtime/renderer:44/44 passed.
- Public style docs/demo:8/8 passed.
- npm run typecheck and npm run lint:both workspaces passed.
- npm test:666/666 passed; /private/tmp/scrollbar-fix-node-tests.log.
- Import audit:150 modules,1118 exports,54 observations, no stale/missing/CommonJS emitted modules; /private/tmp/scrollbar-fix-import-graph.json. Regenerated tracked initial audit evidence restored byte-for-byte from HEAD and remains unchanged.

# ISSUES

Root owns fresh final6px library/app builds, real Edge/Chrome browser runs, independent review and final verify/verify:package gates. Browser runners were extended but not executed by this agent. Authenticated SharePoint validation NOT EXECUTED. No branch changes, agents, commits, pushes or worktrees.
