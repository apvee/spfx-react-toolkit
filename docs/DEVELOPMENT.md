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
| `npm run verify:examples` | Demo registry matches 44 hook and 4 provider exports |
| `npm run verify:runtime-store` | Runtime store behavior check using the local compiler |
| `npm run verify:public-docs` | Public helper/service coverage and local documentation link targets |
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

## Shared Fluent packages

The library declares migration-v8-v9 ^9.9.12 and react-theme ^9.2.0 as mandatory peers. They are devDependencies for library development and dependencies of the SPFx test app. The published package does not own separate Fluent runtime dependencies. The tarball consumer check verifies that the app and library resolve the same Fluent packages. This is an explicit consumer contract selected by the user; all historical peer ranges remain unchanged. `tslib` is supplied by the SPFx packages that declare and import it, rather than an unused direct dependency of this library.
