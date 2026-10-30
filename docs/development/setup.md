# Repository setup

[Development index](INDEX.md). All commands run from the cloned repository root.

Use the [documented Node/toolchain baseline](../getting-started/installation.md#compatibility). To install and build the checkout:

```bash
npm ci
npm run build
```

The [verification guide](verification.md) owns the full command map and explains which checks apply to an ordinary change, a package boundary or a release.

## Layout and build boundary

| Path | Responsibility |
|------|----------------|
| `packages/spfx-react-toolkit/src` | Public library implementation |
| `packages/spfx-react-toolkit/lib` | TypeScript ESNext output, declarations and maps |
| `apps/spfx-react-toolkit-test/src` | WebPart, Application Customizer and demo UI |
| `apps/spfx-react-toolkit-test/config` | SPFx Gulp and tenant debug configuration |
| `apps/spfx-react-toolkit-test/lib`, `dist`, `temp`, `sharepoint` | Generated sample/build/package output under the app workspace |
| `tests`, `scripts`, `docs` | Root behavioral checks, verification and public documentation |
| `.docs/maintenance`, `.docs/superpowers` | Local maintenance records, evidence and plans; existing tracked historical records are retained |

The app resolves `@apvee/spfx-react-toolkit` as an npm workspace package through its compiled `lib/index.js`. It does not compile the library source through a path alias. Build library first; after editing library source, rebuild it and stop/restart the running app serve process. The app's Sass and webpack steps are independent of library compilation.

The old root `src`, `config`, `gulpfile.js` and SPFx output locations moved into their respective workspaces. Consumers retain the same package imports and public exports.

## Local SPFx debug

```bash
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve
```

`npm run serve` rebuilds the library before starting the app. For example, `npm run serve -- --config=spFxReactToolkitTest` forwards the configuration to Gulp.

Trusting the certificate affects your local development certificate store. Configure the app's tenant URLs before opening the authenticated workbench. See [SharePoint validation setup](sharepoint-validation/setup.md) for exact host entry points and manual checks. The repository VS Code launch configuration points at the app workspace; its TypeScript SDK points at root `node_modules`.
