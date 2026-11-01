# Develop and maintain SPFx React Toolkit

Use these procedures to work on the public library, its private SPFx sample, verification tools and documentation. Local tests and browser fixtures provide different evidence from an authenticated SharePoint run.

[Documentation home](../INDEX.md)

| Task | Procedure |
| --- | --- |
| Install, build and debug the checkout | [Repository setup](setup.md) |
| Select checks and inspect rendered behavior | [Verification](verification.md) |
| Author source-owned runnable examples and explore scenarios | [Executable authoring](documentation.md#executable-example-authoring) · [Explorer](setup.md#explore-source-owned-scenarios) |
| Change documentation and reconstruct public outputs | [Documentation maintenance](documentation.md) |
| Validate actual host, permissions and persistence | [Essential SharePoint checklist](sharepoint-validation/INDEX.md) |

The [library workspace](../../../packages/spfx-react-toolkit) owns public source and compiled JavaScript/declarations. The [sample workspace](../../../apps/spfx-react-toolkit-test) owns SPFx entry points, demo UI and host configuration. The sample consumes compiled `lib`: build the library before the app, then rebuild and restart an active app serve process after library source changes.

[AGENTS.md](../../../AGENTS.md) governs branch selection, change scope and completion evidence. Public capability changes require aligned owning JSDoc, generated Markdown and an observable test-web-part scenario. Start application integration with [installation](../getting-started/installation.md), and follow [control slots](../guides/control-slots.md) and [localization](../guides/localization.md) when those boundaries change. Record actual commands and observations; a build or documented procedure does not establish tenant acceptance.
