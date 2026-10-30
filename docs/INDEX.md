# SPFx React Toolkit documentation

SPFx React Toolkit provides React providers, hooks, typed styles, helpers and services for SharePoint Framework applications. It brings together access to SPFx context, component properties, HTTP clients, PnPjs, themes, permissions and storage. Host-specific providers keep runtime state scoped to each SPFx instance.

Use this documentation to integrate the toolkit into an application, choose APIs for a feature, or contribute to the library. If you already know a symbol's name, go directly to the [alphabetical API catalog](api/INDEX.md).

## Start with your first integration

Follow these guides when adding the toolkit to an existing SPFx project:

1. [Install the package and compatible peers](getting-started/installation.md). Check the host's dependencies, shared React/Fluent/Griffel runtimes and the distinction between declared compatibility and the verified toolchain baseline.
2. [Build a first web part](getting-started/quick-start.md). Follow a complete example that mounts a provider, reads user information through a hook and unmounts the React tree when the web part is disposed.
3. [Choose the provider for your host](getting-started/choosing-a-provider.md). Understand the integration points for web parts, Application Customizers, Field Customizers and ListView Command Sets, including their validation limits.

## Browse API families

The public API is organized by the responsibility of each part:

| Family | What it provides | When to use it |
| --- | --- | --- |
| [Providers and core types](api/INDEX.md#providers-and-core-types) | Host integration, instance context and shared TypeScript contracts | Connect a React tree to its actual SPFx component instance. |
| [Hooks](api/INDEX.md#react-and-spfx-hooks) | React access to context, properties, clients, themes, storage and diagnostics | Read toolkit state or invoke operations from a React component. SPFx hooks require a matching provider; generic React utilities document their own dependencies. |
| [Helpers and styles](api/INDEX.md#helpers-and-styles) | Input-based transformations, permission utilities, path builders, serialization and typed style descriptors | Reuse a focused transformation or compose styles. The `useSx` hook uses the host renderer and documented theme requirements. |
| [Services](api/INDEX.md#non-react-services) | SharePoint, PnPjs and Graph operations using caller-supplied clients | Perform reusable I/O in command handlers or other code outside React. |

The [API catalog](api/INDEX.md) lists exported symbols and their import paths, including associated public types. Follow a symbol to its reference for signatures, parameters, defaults, results, error behavior and limitations. Related types stay with their capability; shared core types have a [dedicated reference](api/core/types.md).

## Task guides

These guides explain how to select and combine APIs for a concrete task. Detailed contracts and complete option definitions remain in the linked API references.

| Task | Guide and coverage |
| --- | --- |
| Call an external API, SharePoint REST or Microsoft Graph | [HTTP clients](guides/http-clients.md): client selection, request state, readiness and overlapping invocations. |
| Persist data with the appropriate lifetime and sharing scope | [Storage](guides/storage.md): browser, OneDrive, tenant and site collection options, permissions and failure handling. |
| Check SharePoint permissions or delegated API scopes | [Permissions](guides/permissions.md): choosing a check, interpreting precheck results and handling actual request authorization. |
| Apply typed layout, spacing and theme styles | [Styling](guides/styling.md): composing descriptors with `useSx` and using the shared host runtime. |
| Reuse toolkit operations outside a React component | [Helpers and services](guides/helpers-and-services.md): dependency ownership, examples and explicit PnP feature registrations. |
| Select package entry points and understand bundling | [Package imports](guides/package-imports.md): supported aliases, resolver requirements, tree shaking and compatibility boundaries. |

## Develop and verify the toolkit

The [development documentation](development/INDEX.md) covers the npm workspaces, local setup, build/debug workflow, browser fixtures and verification commands. The private SPFx sample consumes the compiled library, so library source changes require a library rebuild and a restart of the app's serve process.

Use the [SharePoint validation procedures](development/sharepoint-validation/INDEX.md) when checking behavior in an authenticated host. They describe setup, scenario checklists and how to record results. Local tests and builds do not establish tenant permissions, Graph consent or actual SharePoint-host behavior.

For documentation changes, follow the [authoring and maintenance guide](development/documentation.md). It explains content ownership, API page structure, generated catalog/README updates and link verification. Source JSDoc supports IntelliSense, while the sample provides observable usage scenarios alongside the written references.
