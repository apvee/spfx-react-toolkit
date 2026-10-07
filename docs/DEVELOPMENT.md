# Development

All commands below run from the cloned repository root. Use Node `>=22.14.0 <23.0.0`; the lockfile pins SPFx 1.21.1, React/ReactDOM 17.0.1, TypeScript 5.3.3 and PnPjs 4.17.0. The published SPFx peer range `>=1.18.0 <2.0.0` is not a tested version matrix.

## Layout and build boundary

| Path | Responsibility |
|------|----------------|
| `packages/spfx-react-toolkit/src` | Public library implementation |
| `packages/spfx-react-toolkit/lib` | TypeScript ESNext output, declarations and maps |
| `apps/spfx-react-toolkit-test/src` | WebPart, Application Customizer and demo UI |
| `apps/spfx-react-toolkit-test/config` | SPFx Gulp and tenant debug configuration |
| `apps/spfx-react-toolkit-test/lib`, `dist`, `temp`, `sharepoint` | Generated sample/build/package output under the app workspace |
| `tests`, `scripts`, `docs` | Root behavioral checks, verification and public documentation |
| `.docs/maintenance`, `.docs/superpowers` | Internal versioned maintenance records and local ignored agent plans |

The app resolves `@apvee/spfx-react-toolkit` as an npm workspace package through its compiled `lib/index.js`. It does not compile the library source through a path alias. Build library first; after editing library source, rebuild it and stop/restart the running app serve process. The app's Sass and webpack steps are independent of library compilation.

The old root `src`, `config`, `gulpfile.js` and SPFx output locations moved into their respective workspaces. Consumers retain the same package imports and public exports.

## Commands

```bash
npm ci
npm run build:library
npm run build:app
```

`npm ci` uses the checked-in root lockfile. Review installation output: npm may apply its own lifecycle-script policy, so successful installation alone does not prove every dependency install script ran.

| Root command | Actual task |
|--------------|-------------|
| `npm run build:library` | Library workspace `tsc -p tsconfig.json` |
| `npm run build:app` | App workspace `gulp bundle` |
| `npm run build` | Builds library, then app |
| `npm run serve` | Builds library, then starts the app workspace with `gulp serve`; forwards arguments after `--` |
| `npm run clean` | Library removes `lib`; app runs `gulp clean` |
| `npm test` | `node --test tests/*.test.cjs` |
| `npm run typecheck` | Both workspaces run `tsc --noEmit` |
| `npm run lint` | Both workspaces run ESLint on source |
| `npm run verify:examples` | Demo registry matches 46 hook and 4 provider exports plus all style namespaces/scopes |
| `npm run verify:runtime-store` | Runtime store behavior check using the local compiler |
| `npm run verify:public-docs` | Qualified public style exports/JSDoc, historical helper/service coverage and local documentation link targets |
| `npm run verify:api` | Public export/declaration compatibility checks |
| `npm run verify` | Tests, typecheck, lint, example/runtime/docs/API checks |
| `npm run verify:package` | Tarball content and consumer verification |
| `npm run pack:library` | Builds library and runs npm pack into root `artifacts` |
| `npm run bundle:ship` | Builds library, then app `gulp bundle --ship` |
| `npm run package:solution` | Ship bundle, then app `gulp package-solution --ship` |

Run `npm run verify` and `npm run verify:package` before requesting review. `prepublishOnly` is a lifecycle guard, not a publishing instruction: root runs clean/build/verify; the library workspace has its own clean/build/lint/test guard. This workflow does not authorize npm publication, pushes or tenant deployment.

## Local SPFx debug

```bash
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve
```

`npm run serve` rebuilds the library before starting the app. For example, `npm run serve -- --config=spFxReactToolkitTest` forwards the configuration to Gulp.

Trusting the certificate affects your local development certificate store. Configure the app's tenant URLs before opening the authenticated workbench. See [SharePoint validation](./SHAREPOINT-VALIDATION.md) for exact host entry points and manual checks. The repository VS Code launch configuration points at the app workspace; its TypeScript SDK points at root `node_modules`.

## Package shape

`npm run pack:library` writes `artifacts/apvee-spfx-react-toolkit-2.1.0.tgz` for the current package version. The tarball contains package metadata, README, LICENSE and compiled `lib/index.*`, `lib/core`, `lib/hooks`, `lib/helpers`, `lib/services`, `lib/utils`. Entry points remain `lib/index.js` and `lib/index.d.ts`. The app, SPFx manifests, source, tests, config and repository docs are excluded. Source maps point back to source paths; source is not embedded in the tarball.

The isolated consumer probe covers all six public site-store symbols, the hook/service deep imports and existing tenant APIs. See [site storage](./api/hooks/storage.md#usespfxsitekeyvaluestore) and the [standalone service](./api/services/INDEX.md#createspfxsitekeyvaluestoreservice) for their contracts; the [real-host checklist](./SHAREPOINT-VALIDATION.md#site-collection-key-value-store) remains separate.

The list-selector consumer probe covers title strings, `SPFxPnPListSelector`, all three new hook wrappers, generic item types and root/deep imports. Local service/hook tests cover lazy validation, supplied-client targeting and request escaping. Follow the [authenticated list-selector checklist](./SHAREPOINT-VALIDATION.md#list-selector-modes) for server behavior, permissions and special characters.

A package-content check does not validate tenant authentication, Graph consent or SharePoint-host lifecycle. Local regression tests use real React lifecycle with doubles at unavailable SDK/service boundaries. Record real-host results separately.

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

## Shared Fluent packages

The library declares `@griffel/core` (`^1.19.2`), `@griffel/react` (`^1.5.30`), `@fluentui/react-shared-contexts` (`^9.25.2`), `@fluentui/react-migration-v8-v9` (`^9.9.12`), `@fluentui/react-theme` (`^9.2.0`) and `@fluentui/react-utilities` (`^9.25.1`) as mandatory peers. They are devDependencies for library development and dependencies of the SPFx test app. The published package does not own separate Fluent runtime dependencies. The tarball consumer check verifies that the app and library resolve the same Fluent packages. These are explicit shared consumer contracts; all historical peer ranges remain unchanged. `useStableCallback` is a direct alias of `useEventCallback` from react-utilities, not a separate implementation. `tslib` is supplied by the SPFx packages that declare and import it, rather than an unused direct dependency of this library.

## Styles sample and documentation checks

The lazy Styles panel imports only package-root exports. Its catalog controls expose every namespace and base selection, with dedicated width/gap/column actions, theme and preset choices, conditional enabled states, two independent named query regions, viewport scopes and a constrained native scroll region. React Hooks/useStableCallback remains its own scenario.

Run `node --test tests/public-style-docs.test.cjs tests/demo-sx.test.cjs`, `npm run verify:public-docs` and `npm run verify:examples`. The documentation checker follows directory `index.ts` barrels, named re-exports and `export * as` namespaces recursively, including `overflow.horizontal` and `overflow.vertical`. It checks each qualified public member and its source JSDoc, plus public interface members. Private descriptor brands and internal engine files are not root exports and do not become public documentation obligations. Historical helper/service inventory checks stay enabled. The example checker requires the registry and interactive sample source to cover every style namespace/scope alongside the existing hook/provider checks.

Build library before app with `npm run build:library`, then `npm run build:app`; restart an active serve before evaluating new compiled library behavior. Record the actual commands, exit codes and Gulp warnings, even when the outer build exits successfully. The sample uses optional `@fluentui/react-provider@9.22.8` with toolkit themes; consumers are not required to add that provider solely for useSx.

Local Node/DOM checks and the standalone browser fixture establish their own behavior only. The authenticated Styles checklist in [SharePoint validation](./SHAREPOINT-VALIDATION.md#styles-and-usesx) is **NOT EXECUTED** until a tenant run records it. Custom themes, unsupported inverted high contrast, native scrollbar differences and arbitrary-value CSS growth remain documented in the [style reference](./api/helpers/styles.md).
