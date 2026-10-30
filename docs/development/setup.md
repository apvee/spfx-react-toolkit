# Set up and run the repository

[Documentation home](../INDEX.md) · [Development](INDEX.md)

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
| `.docs/maintenance`, `.docs/superpowers` | Local Git-ignored maintenance records, evidence and agent plans |

The app resolves `@apvee/spfx-react-toolkit` as an npm workspace package through its compiled `lib/index.js`. It does not compile the library source through a path alias. Build library first; after editing library source, rebuild it and stop/restart the running app serve process. The app's Sass and webpack steps are independent of library compilation.

The old root `src`, `config`, `gulpfile.js` and SPFx output locations moved into their respective workspaces. Consumers retain the same package imports and public exports.

## Authoring control presentation

New or extended named presentation regions follow [Control presentation slots](../guides/control-slots.md). Use the private Fluent resolver/types with fixed or documented bounded tags, preserve shorthand DOM and semantic shells, and keep structural children, domain callbacks and inherited native slots feature-specific. JSDoc and each control guide must describe inner ref targets, callback forwarding, formatting and omission/default behavior. Existing `action`/`actions` names remain stable.

Use host-rendered classes for toolkit styles. The private resolver accepts `useSx` or an existing Griffel string class composer where import boundaries forbid the SX engine; it creates no renderer. Consumer style is forwarded, and required refs/events retain their feature contracts. See [ContentCard](../api/controls/content-card.md) and [PreviewCard](../api/controls/preview-card.md) for native Card composition, explicit caller controls and preview precedence.

## Commands

```bash
npm ci
npm run build:library
npm run build:app
```

`npm ci` uses the checked-in root lockfile. Review installation output: npm may apply its own lifecycle-script policy, so successful installation alone does not prove every dependency install script ran.

| Root command | Actual task |
|--------------|-------------|
| `npm run build:library` | Validates localization inventory, runs library `tsc -p tsconfig.json`, then emits AMD resource files |
| `npm run build:app` | App workspace `gulp bundle` |
| `npm run build` | Builds library, then app |
| `npm run serve` | Builds library, then starts the app workspace with `gulp serve`; forwards arguments after `--` |
| `npm run clean` | Library removes `lib`; app runs `gulp clean` |
| `npm test` | `node --test tests/*.test.cjs` |
| `npm run typecheck` | Both workspaces run `tsc --noEmit` |
| `npm run lint` | Both workspaces run ESLint on source |
| `npm run verify:examples` | Demo registry matches 47 hook and 4 provider exports plus public controls, the complete Form compound family and all style namespaces/scopes |
| `npm run verify:runtime-store` | Runtime store behavior check using the local compiler |
| `npm run verify:public-docs` | Public control/style surfaces, selected source JSDoc/member coverage, helper/service inventories and local documentation link targets |
| `npm run verify:api` | Public export/declaration compatibility checks |
| `npm run verify:import-graph` | Audits public/internal imports and module evaluation boundaries |
| `npm run verify:bundle-contract -- --consumer-root <directory>` | Runs the production bundle gate against an installed-package consumer |
| `npm run verify:localization` | Checks the resource inventory and localization/consumer-CLI lifecycle boundaries |
| `npm run verify` | Tests, typecheck, lint, example/runtime/docs/API and import-graph checks |
| `npm run verify:package` | Fresh tarball consumer resolution, bundle/runtime/mutation gates and SPFx verification |
| `npm run pack:library` | Builds library and runs npm pack into root `artifacts` |
| `npm run bundle:ship` | Builds library, then app `gulp bundle --ship` |
| `npm run package:solution` | Ship bundle, then app `gulp package-solution --ship` |

## GitHub Actions

The `Verify monorepo` workflow in `.github/workflows/verify.yml` runs only through `workflow_dispatch`. Pushes and pull request updates do not start it automatically. To run the complete installation, build, verification and package checks, open **Actions > Verify monorepo > Run workflow**, select the branch to verify and start the run.

[GitHub requires the workflow with its manual trigger to be present on the repository's default branch](https://docs.github.com/en/actions/how-tos/manage-workflow-runs/manually-run-a-workflow) before manual dispatch is available. Merging this change into another branch alone does not enable the button. Changing the triggers does not cancel runs that have already started.

## Local SPFx debug

```bash
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve
```

`npm run serve` rebuilds the library before starting the app. For example, `npm run serve -- --config=spFxReactToolkitTest` forwards the configuration to Gulp.

Trusting the certificate affects your local development certificate store. Configure the app's tenant URLs before opening the authenticated workbench. See [SharePoint validation](sharepoint-validation/INDEX.md) for exact host entry points and manual checks. The repository VS Code launch configuration points at the app workspace; its TypeScript SDK points at root `node_modules`.

### Hosted workbench bootstrap and debug culture

The hosted workbench can attempt automatic manifest discovery on the default port 4321 while your configured debug manifest loads on another port. An error naming that fallback alone is not the readiness result for the configured manifest. After **Load debug scripts**, dismiss the fallback warning if present, open the web part toolbox, confirm **SPFxReactToolkitTest**, insert it and wait for its tabs to render. An error for the configured manifest, a missing toolbox entry after registration, or a failing component load requires separate investigation; do not ignore arbitrary loader errors or bypass certificate warnings.

For native Italian resources, append `locale=it-it` to the actual hosted workbench URL, for example:

```text
https://apveelabs.sharepoint.com/sites/apvee-content-flow-translator-demo/_layouts/15/workbench.aspx?debug=true&noredir=true&debugManifestsFile=https%3A%2F%2Flocalhost%3A4322%2Ftemp%2Fbuild%2Fmanifests.js&locale=it-it
```

The generated SPFx debug manifest reads this page parameter. It is distinct from per-instance React `strings` and does not change account language settings. Verify actual resource requests and UI versus regional culture using the [localization host checklist](sharepoint-validation/checklists/localization.md).

## Shared Fluent packages

The library declares `@griffel/core` (`^1.21.4`), `@griffel/react` (`^1.7.8`), `@fluentui/react-shared-contexts` (`^9.26.4`), `@fluentui/react-migration-v8-v9` (`^9.10.20`), `@fluentui/react-theme` (`^9.2.2`), `@fluentui/react-components` (`^9.74.9`) `@fluentui/react-utilities` (`^9.26.7`) and `@fluentui/react-charts` (`^9.3.27`) as mandatory peers. They are devDependencies for library development and dependencies of the SPFx test app. The published package does not own separate Fluent runtime dependencies. FormattedText uses library runtime dependencies unified 11.0.5, remark-parse 11.0.0 and remark-gfm 4.0.1. The tarball consumer check verifies that the app and library resolve the same Fluent packages. These are explicit shared consumer contracts; SPFx, React and PnP peer ranges remain unchanged; the tested Fluent/Griffel floors are updated together. `useStableCallback` is a direct alias of `useEventCallback` from react-utilities, not a separate implementation. `tslib` is supplied by the SPFx packages that declare and import it, rather than an unused direct dependency of this library.
