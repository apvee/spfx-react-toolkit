# SPFx React Toolkit documentation

SPFx React Toolkit provides React controls, providers, hooks, typed styles, helpers and services for SharePoint Framework applications. It brings together access to SPFx context, component properties, HTTP clients, PnPjs, themes, permissions and storage. Host-specific providers keep runtime state scoped to each SPFx instance.

Use this documentation to integrate the toolkit into an application, choose APIs for a feature, or contribute to the library. If you already know a symbol's name, go directly to the [alphabetical API catalog](../api/INDEX.md).

## Start with your first integration

Follow these guides when adding the toolkit to an existing SPFx project:

1. [Install the package and compatible peers](getting-started/installation.md). Check the host's dependencies, shared React/Fluent/Griffel runtimes and the distinction between declared compatibility and the verified toolchain baseline.
2. [Build a first web part](getting-started/quick-start.md). Follow a complete example that mounts a provider, reads user information through a hook and unmounts the React tree when the web part is disposed.
3. [Choose the provider for your host](getting-started/choosing-a-provider.md). Understand the integration points for web parts, Application Customizers, Field Customizers and ListView Command Sets, including their validation limits.

## Task guides

Controls compose interactions; native Property Pane factories configure the host. Providers connect an actual SPFx instance, hooks expose React-facing state and operations, helpers/styles transform explicit inputs, and services perform I/O with caller-supplied clients outside React. Associated types stay with their capability; the generated catalog discovers their legal imports.

Choose a task guide to understand integration order, application-owned state and verification boundaries. Follow its canonical API links for signatures, defaults, legal imports and associated types. Navigation below is discovered from the source files.

## Develop and verify the toolkit

The [development documentation](development/INDEX.md) covers the npm workspaces, local setup, build/debug workflow, browser fixtures and verification commands. The private SPFx sample consumes the compiled library, so library source changes require a library rebuild and a restart of the app's serve process.

Use the [SharePoint validation procedures](development/sharepoint-validation/INDEX.md) when checking behavior in an authenticated host. They describe setup, one essential boundary checklist and how to record results. Local tests and builds do not establish tenant permissions, Graph consent or actual SharePoint-host behavior.

For documentation changes, follow the [authoring and maintenance guide](development/documentation.md). It explains content ownership, API page structure, generated catalog/README updates and link verification. Source JSDoc supports IntelliSense, while the sample provides observable usage scenarios alongside the written references.
