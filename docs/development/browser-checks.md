# Run local browser checks

[Documentation home](../INDEX.md) · [Development](INDEX.md)

## Local useSx browser checkpoint

The standalone fixture in `tests/fixtures/sx-browser/` imports the generic source modules with real React 17, Griffel and FluentProvider. It exercises the public catalog and independent query/state cases; private diagnostic recipes cover compiler boundaries. Fixture host styles supply geometry; `sx` supplies class strings. This is a local browser protocol, independent of the SPFx sample and authenticated SharePoint validation.

Build with the installed Webpack and TypeScript toolchain, then serve the generated fixture:

```bash
node scripts/build-sx-browser-fixture.cjs
node scripts/serve-sx-browser-fixture.cjs
```

The default URL is `http://127.0.0.1:4317`. Set `SX_BROWSER_PORT` to select another port. The builder typechecks the fixture and its source imports before bundling into ignored `temp/sx-browser/`; it installs no dependencies. Rebuild after generic source edits. This fixture reads source directly; the separate SPFx sample still requires a library build and a serve restart.

Use an available browser connector when present. If it is unavailable, the reproducible fallback script accepts an existing Playwright installation and browser executable:

```bash
PLAYWRIGHT_MODULE_PATH=/path/to/installed/playwright \
SX_BROWSER_EXECUTABLE=/path/to/browser \
node scripts/check-sx-browser-fixture.cjs
```

Omit those variables when normal Node resolution and Playwright's installed browser are available. `SX_BROWSER_URL` overrides the server URL; `SX_BROWSER_EVIDENCE` selects an output directory; `SX_BROWSER_RUN_LABEL` prefixes the JSON, DOM and screenshot files. The default evidence directory is `.docs/maintenance/evidence/use-sx/`. The runner uses Playwright's library directly, with no `@playwright/test` dependency. It launches, checks and closes in one process so it does not rely on a persistent CLI daemon. Restricted environments may require approval to listen on localhost or launch the local browser. Record actual approval outcomes with the evidence.

`expected.json` records the independent computed-style expectations before cascade fixes. Both fresh-page mount orders test inclusive container thresholds at 480/640/1024px, nearest nested and independent containers, vertical inline size, self-container exclusion, unchanged explicit viewport queries, omitted color inheritance, parent scope isolation, removed scopes after rerender, separate `sx` strings and external Griffel classes in both orders. Pointer and keyboard events test hover, active and focus-visible together, container/media plus state, each state's responsive priority, enabled/disabled and selected changes under a stationary pointer, native keyboard focus outlines and LTR/RTL logical/physical properties. A mobile viewport supplements the desktop checks.

Inspect the JSON observations, failure logs, DOM captures and screenshots. The runner checks page identity, content, framework overlays and every console warning/error; it serves an empty favicon response to avoid hiding missing assets behind an error filter. Browser screenshots retain native scrollbar visibility, although native appearance depends on the browser and operating system. Computed styles, real pointer/keyboard events and visible focus indicators establish this local checkpoint; CSSOM/JSDOM assertions alone do not. Record authenticated SharePoint-host validation separately as executed or not executed.
