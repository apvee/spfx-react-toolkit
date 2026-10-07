# Actual ImportsPanel local Chrome QA

## RESULT

Executed actual sample ImportsPanel in installed Chrome 155.0.8059.39 using actual compiled package root/domain/legacy entry points. No toolkit alias, hook replacement, or SDK runtime fixture was used. 62 observations recorded; one failed check: mobile override horizontal overflow. No product files changed by this QA worker.

## FINDING

At viewport 390×844, click Root imports → Apply width override. The preview has computed content width 360px plus logical padding 16px; its left edge is 33px and right edge 409px. Page scrollWidth becomes 409px against innerWidth 390px. Text extends beyond the card. Initial 240px mobile state and desktop 1440×1000 pass overflow inspection. See `mobile-override.png` and `results.json` mobile override overflow inspection. This is a sample layout finding; no production change was attempted.

## EVIDENCE

- Actual three cards render, page identity/origin correct, meaningful content and no framework error overlay.
- Independent counters 1/2/3 persist across interactions. Original captured callbacks observe current state and report same identity.
- Root/domain/legacy computed typography, width and padding match; active class strings match. Overrides produce 360px equivalently, removal restores 240px and original classes.
- FluentProvider RTL produces paddingRight 16px / paddingLeft 0px across all three previews; LTR restores paddingLeft 16px / paddingRight 0px.
- All mobile increment controls remain visible. Desktop and initial mobile layouts have no overflowing elements.
- Zero browser console warnings/errors, page errors, or failed requests; console events retained in results JSON.
- Host icon initialization uses real `initializeIcons('/fonts/')` and installed Fluent V8 font assets served locally, without warning suppression.
- Desktop initial/overrides/RTL and mobile initial/override full-page screenshots preserved. Mobile override and desktop RTL were visually inspected by QA worker.

## ENVIRONMENT AND REPRODUCTION

Browser plugin not available; Playwright fallback uses existing cached library at `/Users/fabiofranzini/.npm/_npx/e41f203b7505f1fb/node_modules/playwright` and `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`. No installation or upgrade.

Command from repository root:

```sh
node /private/tmp/imports-panel-browser-qa/run.cjs
```

Preserved script copy: `run.cjs` in this evidence directory. To restore temp script, copy it to `/private/tmp/imports-panel-browser-qa/run.cjs`. Script creates fixture files under that directory, uses an ephemeral unique localhost port, and closes only its own Chrome and HTTP server. This environment required sandbox escalation for localhost listening/Chrome; the approval succeeded.

Webpack mode production, minimize false, usedExports true, sideEffects true; actual sample TS transpiled with repository `scripts/sx-browser-typescript-loader.cjs`; resolution uses installed workspace dependencies and public package exports. SDK `@microsoft/sp-*` requests are CommonJS host externals, following repository bundle-contract convention. No runtime externals are supplied. Emitted bundle assertions confirm no SPFx host require or SDK implementation markers retained. Webpack module inventory is preserved in `webpack-modules.json`.

The first build without host externals attempted to parse unrelated SDK `.resx`/SCSS and missing private `@ms` packages, before tree shaking could remove them. Initial error retained in `initial-sdk-build-failure.log`; final successful webpack log in `build.log`. Existing baseline-browser-mapping age warning remains unsuppressed in `run.log`; this is a build-tool warning, not a browser runtime console warning. No dependency update attempted.

## ISSUES

Mobile override overflow remains unresolved pending root review. Authenticated SPFx/SharePoint host execution: NOT EXECUTED. This fixture proves actual panel/React17/Fluent/Griffel behavior in a generic local browser host only. It does not establish tenant permissions or SPFx host behavior.
