# Set up and run the repository

Install from the root lockfile, build the public library before the private sample, and debug against the intended SPFx host. The [installation guide](../getting-started/installation.md) distinguishes declared peer compatibility from the verified toolchain baseline; the root [package metadata](../../../package.json) owns the supported Node range.

[Documentation home](../INDEX.md) · [Development](INDEX.md)

## Layout and build boundary

Public implementation lives in `packages/spfx-react-toolkit/src`; its `lib` directory is compiler output. The sample's source, Gulp configuration and generated `lib`, `dist`, `temp` and `sharepoint` outputs live under `apps/spfx-react-toolkit-test`. Root `tests` and `scripts` own behavioral checks and verification tooling. Reusable fixtures are versioned in `tests/fixtures`; local maintenance evidence and plans belong in ignored `.docs/maintenance` and `.docs/superpowers`.

```bash
npm ci
npm run build:library
npm run build:app
```

Review install output: successful installation alone does not prove every dependency lifecycle script ran. The [library build script](../../../packages/spfx-react-toolkit/package.json) checks localization inventory, compiles TypeScript, emits AMD resources. It does not generate or check documentation or executable examples. A complete checkout uses the versioned generated outputs. The [app build](../../../apps/spfx-react-toolkit-test/package.json) runs `gulp bundle`; it consumes the workspace package's compiled entry, without a source path alias. After editing JSDoc, API or editorial sources, run `npm run generated:sync`, review and version the source/output diff together, then run `npm run generated:check` and relevant build/tests. After editing library source, keep the order sync → rebuild library → stop/restart a running app serve before evaluating behavior. Build success and output freshness are separate proofs; no generation/check hook or JSDoc watcher runs during build or startup. `npm run build` orders both builds; root [scripts](../../../package.json) own ship bundling, packaging, cleanup and verification commands.

## Local SPFx debug

```bash
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve
```

Trusting the development certificate changes the local certificate store. Root serve rebuilds the library, then forwards arguments to app `gulp serve`; use `npm run serve -- --config=spFxReactToolkitTest` for the named configuration. Configure tenant URLs as described in [authenticated setup](sharepoint-validation/INDEX.md#prepare-a-validation-run). The checked-in [serve configuration](../../../apps/spfx-react-toolkit-test/config/serve.json) uses HTTPS on port 4322. The [VS Code launch configuration](../../../.vscode/launch.json) uses the app workspace as webRoot; [editor settings](../../../.vscode/settings.json) use root `node_modules/typescript/lib`.

### Hosted workbench bootstrap and debug culture

The hosted workbench may attempt default manifest discovery on port 4321 while the configured manifest uses 4322. A warning naming that fallback alone is not the readiness result for the configured manifest. After **Load debug scripts**, dismiss the fallback warning if present, open the toolbox, confirm **SPFxReactToolkitTest**, insert it and wait for the example explorer to render. Investigate errors for the configured manifest, a missing toolbox entry or component load failure; do not bypass certificate warnings or ignore arbitrary loader errors. The [sample Gulp integration](../../../apps/spfx-react-toolkit-test/gulpfile.js) sets the configured origin on debug cumulative manifests for the pinned SPFx build.

Append `locale=it-it` to the actual hosted workbench URL to request native Italian resources, keeping its existing debug parameters. The generated debug manifest reads this page parameter; it is distinct from per-instance React `strings`, regional formatting culture and account language. Observe actual resource requests and UI as described in [localization and instance isolation](sharepoint-validation/INDEX.md#localization-and-instance-isolation).

## Shared runtimes and Keyborg integration

The consuming app supplies the library's mandatory shared peers. Treat [library peer metadata](../../../packages/spfx-react-toolkit/package.json), the [app dependency set](../../../apps/spfx-react-toolkit-test/package.json) and [lockfile](../../../package-lock.json) as owners; do not install an independent Fluent/Griffel/React runtime. Provider packages may differ in version while still resolving the same shared contexts, theme and renderer. Package checks verify actual resolution rather than inferring identity from version labels.

The coordinated Fluent/Charts navigation subtree locks Tabster 8.8.0 with Keyborg 2.14.1. Tabster 8.8.1 resolves Keyborg 3 alongside Fluent's Keyborg 2, and the actual React 17 cleanup probe detects a registry collision. Preserve the reviewed lock selection; it satisfies the published range without overrides.

The sample [Gulp configuration](../../../apps/spfx-react-toolkit-test/gulpfile.js) applies [keyborg-instance-id-loader.cjs](../../../apps/spfx-react-toolkit-test/config/keyborg-instance-id-loader.cjs) to the installed ESM/CJS entries. Independently bundled SharePoint and solution copies share `window.__keyborg`, but Keyborg 2.14.1's counter is module-local. The loader adds a 128-bit random suffix using browser `crypto.getRandomValues` so bundles do not overwrite each other's references. Native subscriptions and cleanup remain intact; no registry is added or diagnostic suppressed. The loader fails when its reviewed declaration changes. Run `node --test tests/spfx-keyborg.test.cjs tests/spfx-debug-port.test.cjs` when changing this integration. This workaround belongs to the sample build; it does not rewrite published library output or other consumers' configurations.

## GitHub Actions

The [Verify monorepo workflow](../../../.github/workflows/verify.yml) uses `workflow_dispatch` only. In **Actions > Verify monorepo > Run workflow**, select the branch and start the run when authorized. The manual trigger must exist on the default branch for dispatch to be available. Push/PR updates do not start it automatically, and trigger edits do not cancel existing runs. CI checks generated documentation and executable examples explicitly in read-only mode before the build; see [documentation maintenance](documentation.md#deterministic-reconstruction-and-commands). Running local checks does not authorize publication, deployment or tenant approvals.

## Explore source-owned scenarios

Open a family group and select a feature, then choose its **Scenario** in the central panel. Groups expand independently; collapsing the active group keeps its content and identifies the selected feature. **Ripristina scenario** restores that scenario's baseline and invalidates its pending work. It preserves the component; switching scenarios runs cleanup and mounts the selected baseline. Diagnostics start closed. On narrow columns use the **Funzionalità** select with family optgroups; the central Scenario selector remains available. See [executable authoring](documentation.md#executable-example-authoring) for metadata, regions, reset ownership and the optional actual-Shell port.

The focused migration checkpoint covers eleven scenarios across five features: the original nine-case pilot (three Shell, three Pagination, two stable-callback and one explicit list-read scenario), plus advanced four-mode Pagination and root/domain/historical import interoperability owned by `useSx`. The full catalogue additionally retains the independent Shell metadata helper scenario. Source-owned availability, local behavioral verification and authenticated host execution are separate evidence. The read case starts with no list title: choose an existing resource and use its explicit command; selecting the feature or typing a title must not query it. Legacy panels remain in the transitional group until their complete behavior and direct fixture/test consumers have migrated. A passing pilot is not final migration completion.
