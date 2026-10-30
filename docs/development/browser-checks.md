# Local browser and sample checks

[Development index](INDEX.md). All commands run from the cloned repository root.

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

## Styles sample and documentation checks

The lazy Styles panel imports only package-root exports. Its catalog controls expose every namespace and base selection, with dedicated width/gap/column actions, theme and preset choices, conditional enabled states, two independent named query regions, viewport scopes and a constrained native scroll region. React Hooks/useStableCallback remains its own scenario.

Run `node --test tests/public-style-docs.test.cjs tests/demo-sx.test.cjs`, `npm run verify:public-docs` and `npm run verify:examples`. The documentation checker follows directory `index.ts` barrels, named re-exports and `export * as` namespaces recursively, including `overflow.horizontal` and `overflow.vertical`. It checks each qualified public member and its source JSDoc, plus public interface members. Private descriptor brands and internal engine files are not root exports and do not become public documentation obligations. Historical helper/service inventory checks stay enabled. The example checker requires the registry and interactive sample source to cover every style namespace/scope alongside the existing hook/provider checks.

Build library before app with `npm run build:library`, then `npm run build:app`; restart an active serve before evaluating new compiled library behavior. Record the actual commands, exit codes and Gulp warnings, even when the outer build exits successfully. The sample uses optional `@fluentui/react-provider@9.22.8` with toolkit themes; consumers are not required to add that provider solely for useSx.

Local Node/DOM checks and the standalone browser fixture establish their own behavior only. The authenticated Styles checklist in [SharePoint validation](sharepoint-validation/checklists/styling.md#styles-and-usesx) is **NOT EXECUTED** until a tenant run records it. Custom themes, unsupported inverted high contrast, native scrollbar differences and arbitrary-value CSS growth remain documented in the [style reference](../api/helpers/styles.md).

## Imports sample

The lazy `Imports` entry (`spfx-demo-imports`) renders three dedicated stable root/domain/legacy components. Each retains a callback captured at mount and reads the latest counter after incrementing. Width override/removal, body typography and logical start padding compose mixed imports beneath a shared FluentProvider; direction switching exercises the same active context. Existing React Hooks and Styles panels retain their root imports and own their existing export coverage. Imports uses empty symbol coverage plus a dedicated observable-scenario check, preserving 46 hooks and four providers without duplicate symbols.

The app normally imports the root. Only `ImportsPanel.tsx` has the exact approved domain/legacy specifier allowlist; this is not permission for arbitrary app `/lib` imports. Run `node --test tests/demo-imports.test.cjs tests/monorepo.test.cjs`, `npm run verify:examples` and `npm run verify:public-docs`, then build library before app. Rebuild and restart an active serve after library edits. Node/DOM interaction checks verify callback/class behavior; actual local browser computed-style observations and authenticated SharePoint checks must be recorded separately. Build success alone does not execute the scenario.
