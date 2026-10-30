# Installation and compatibility

[Documentation home](../INDEX.md) · [First web part](quick-start.md)

Install into the consuming SPFx application:

```bash
npm install @apvee/spfx-react-toolkit "@fluentui/react-migration-v8-v9@^9.9.12" "@fluentui/react-theme@^9.2.0" "@fluentui/react-utilities@^9.25.1" "@griffel/core@^1.19.2" "@griffel/react@^1.5.30" "@fluentui/react-shared-contexts@^9.25.2"
```

The host supplies React, ReactDOM, SPFx and compatible `@pnp/core`, `@pnp/queryable` and `@pnp/sp` v4 peers. Preserve host-compatible versions and check the resolved dependency tree; modern npm may install missing peers automatically. Install missing peers in the consuming application rather than relying on a repository's development dependencies. This is the compatibility contract on this branch.

## Compatibility

| Requirement | Repository verification baseline |
|-------------|----------------------------------|
| Node.js | `>=22.14.0 <23.0.0` |
| SPFx build/runtime packages | `1.21.1` |
| React / ReactDOM | `17.0.1` (package peers: `17.x`) |
| TypeScript | `5.3.3` |
| PnPjs | `4.17.0` (package peers: `^4.0.0`) |

The package declares SPFx peers `>=1.18.0 <2.0.0`; that range is a compatibility declaration, not evidence of testing every version. Repository checks use SPFx 1.21.1. Authenticated tenant validation is a separate step.

## Shared runtimes

React, Fluent and Griffel instances must be shared with the host. Token-based `useSx` descriptors need theme variables in scope, normally from the host's existing FluentProvider. The hook does not create a renderer or require an SPFx provider. The sample's `@fluentui/react-provider@9.22.8` is optional for consumers. `tslib` is supplied by the SPFx packages that depend on it; the library's ES2020 output does not import it.

See [package imports](../guides/package-imports.md) for bundler/resolver support, legacy imports, PnP registrations and verified bundle boundaries. For a source checkout, use [repository setup](../development/setup.md) instead of treating development scripts as commands available in an installed npm package.
