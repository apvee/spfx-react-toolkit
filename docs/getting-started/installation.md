# Install the toolkit and shared peers

[Documentation home](../INDEX.md) · [Quick start](quick-start.md) · [Localization](../guides/localization.md)

Install in the **consuming SPFx application**:

```bash
npm install @apvee/spfx-react-toolkit
```

The host supplies compatible React, ReactDOM and SPFx runtimes. Keep those versions aligned with the host rather than upgrading its toolchain just to install the toolkit. The package also declares `@pnp/core`, `@pnp/queryable` and `@pnp/sp` v4 peers for its PnPjs APIs.

Add missing shared UI peers using compatible versions. All packages below are mandatory package peers, including Charts and the Form adapters; reuse compatible installations already present in the application:

```bash
npm install \
  "@fluentui/react-components@^9.74.9" \
  "@fluentui/react-charts@^9.3.27" \
  "@fluentui/react-migration-v8-v9@^9.10.20" \
  "@fluentui/react-theme@^9.2.2" \
  "@fluentui/react-utilities@^9.26.7" \
  "@fluentui/react-shared-contexts@^9.26.4" \
  "@griffel/core@^1.21.4" \
  "@griffel/react@^1.7.8" \
  react-hook-form@7.87.0 \
  @fluentui/react-datepicker-compat@0.6.16 \
  @fluentui/react-timepicker-compat@0.4.21 \
  @fluentui/react-icons@2.0.312
```

Modern npm can install missing peers automatically. Review the resolved versions and preserve shared React/Fluent/Griffel identities; see [Form dependencies](../api/controls/form.md#dependencies) for the tested adapter baseline.

### Compatibility

| Dependency | Declared contract | Repository verification baseline |
| --- | --- | --- |
| SPFx runtime peers, including Property Pane | `>=1.18.0 <2.0.0` | `1.21.1` |
| React / ReactDOM | `17.x` | `17.0.1` |
| PnPjs | `^4.0.0` | `4.17.0` |
| Node.js | `>=22.14.0 <23.0.0` | `22.20.0` |
| TypeScript | Host toolchain | `5.3.3` |

The declared SPFx range is broader than the verified baseline; it does not imply testing every SPFx release. Theme-based controls and styles need Fluent theme variables and the shared host renderer.

## Register localized resources

Before using localized controls, configure the consumer postinstall and run the setup CLI once. Follow the [localization guide](../guides/localization.md) for resource mappings, dictionary overrides, monorepo --project setup and the read-only --check command. Complete overrides do not remove the native resource requirement.


## Complete mandatory peer contract

All peers below are declared by the current package; a declared range is not a tested matrix. Reuse the host React/SPFx runtimes and keep shared UI identities aligned.

| Peer | Declared range | Repository version |
| --- | --- | --- |
| `@microsoft/sp-application-base` | `>=1.18.0 <2.0.0` | `1.21.1` |
| `@microsoft/sp-component-base` | `>=1.18.0 <2.0.0` | `1.21.1` |
| `@microsoft/sp-core-library` | `>=1.18.0 <2.0.0` | `1.21.1` |
| `@microsoft/sp-http` | `>=1.18.0 <2.0.0` | `1.21.1` |
| `@microsoft/sp-listview-extensibility` | `>=1.18.0 <2.0.0` | `1.21.1` |
| `@microsoft/sp-page-context` | `>=1.18.0 <2.0.0` | `1.21.1` |
| `@microsoft/sp-webpart-base` | `>=1.18.0 <2.0.0` | `1.21.1` |
| `@pnp/core` | `^4.0.0` | `4.17.0` |
| `@pnp/queryable` | `^4.0.0` | `4.17.0` |
| `@pnp/sp` | `^4.0.0` | `4.17.0` |
| `react` | `17.x` | `17.0.1` |
| `react-dom` | `17.x` | `17.0.1` |
| `@fluentui/react-migration-v8-v9` | `^9.10.20` | `9.10.20` |
| `@fluentui/react-theme` | `^9.2.2` | `9.2.2` |
| `@fluentui/react-utilities` | `^9.26.7` | `9.26.7` |
| `@griffel/core` | `^1.21.4` | `1.21.4` |
| `@griffel/react` | `^1.7.8` | `1.7.8` |
| `@fluentui/react-shared-contexts` | `^9.26.4` | `9.26.4` |
| `@fluentui/react-components` | `^9.74.9` | `9.74.9` |
| `react-hook-form` | `>=7.87.0 <8.0.0` | `7.87.0` |
| `@fluentui/react-datepicker-compat` | `>=0.6.16 <1.0.0` | `0.6.16` |
| `@fluentui/react-timepicker-compat` | `>=0.4.21 <1.0.0` | `0.4.21` |
| `@fluentui/react-icons` | `>=2.0.312 <3.0.0` | `2.0.312` |
| `@fluentui/react-charts` | `^9.3.27` | `9.3.27` |
| `@microsoft/sp-property-pane` | `>=1.18.0 <2.0.0` | `1.21.1` |


## Shared runtimes

React, Fluent and Griffel instances must be shared with the host. Token-based `useSx` descriptors need theme variables in scope, normally from the host's existing FluentProvider. The hook does not create a renderer or require an SPFx provider. The sample's `@fluentui/react-provider@9.22.8` is optional for consumers. `tslib` is supplied by the SPFx packages that depend on it; the library's ES2020 output does not import it.

See [package imports](../guides/package-imports.md) for bundler/resolver support, legacy imports, PnP registrations and verified bundle boundaries. For a source checkout, use [repository setup](../development/setup.md) instead of treating development scripts as commands available in an installed npm package.
