# Root verification — actual ImportsPanel browser

Actual Chrome 155.0.8059.39 run after mobile containment fix: PASS, exit 0, 65 observations with no failures, console events, page errors or failed requests. Command: `node /private/tmp/imports-panel-browser-qa/run.cjs`. Root inspected `results.json`, `mobile-override.png` and `desktop-rtl.png`.

Actual sample panel uses actual compiled toolkit root/domain/legacy entrypoints. Independent captured callbacks see counts1/2/3 and keep identity; equivalent real classes/computed styles override240→360→240; padding16 swaps left/right under FluentProvider LTR/RTL. Desktop1440×1000 and mobile390×844 inspected. The fixed360px preview stays in a named focusable local horizontal scroll region; mobile page scrollWidth375 against viewport390, including override. Focus and ArrowRight scrolling were verified on the actual mobile region while its preview retained computed width360px. Preview itself intentionally extends inside that bounded region.

Initial failing browser evidence is preserved in `red-browser-initial/`; initial SDK parse-stage failure is retained there too. The harness externalizes unused SPFx host modules with the production gate convention, then asserts none emitted. No toolkit mocks, SDK hook replacement, browser warning filters, dependency installations or upgrades. Cached Playwright fallback: Browser plugin not available. Existing build-tool metadata age advisory remains visible in logs.

Generic local browser fixture executed; authenticated SPFx/SharePoint scenario NOT EXECUTED. Only Chrome was checked, not other browsers. The isolated server and browser were closed.
