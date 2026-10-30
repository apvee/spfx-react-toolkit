# Develop and maintain SPFx React Toolkit

[Documentation home](../INDEX.md)

This section is for contributors working on the library, its private SPFx sample, verification tools or documentation. It explains how the checkout is organized, how to run the sample and which evidence each kind of check provides. To install the published package in an application, start with [installation and compatibility](../getting-started/installation.md).

## Choose a development workflow

| Task | Guide | What it covers |
| --- | --- | --- |
| Set up the checkout and run the sample | [Repository setup](setup.md) | Workspaces, build order, local debugging, certificates and SPFx serve configuration. |
| Select checks for a change | [Verification commands](verification.md) | Command responsibilities, scoped verification, package contents and import/bundle gates. |
| Inspect rendered behavior locally | [Browser and sample checks](browser-checks.md) | Standalone browser fixtures, Styles/Imports scenarios and their verification limits. |
| Change public documentation | [Documentation maintenance](documentation.md) | Content ownership, reference structure, generated catalog/README updates and link checks. |
| Exercise behavior in an authenticated host | [SharePoint validation](sharepoint-validation/INDEX.md) | Tenant setup, focused manual checklists, recording results and cleanup. |

## Understand the build boundary

The [library workspace](../../packages/spfx-react-toolkit) owns the public source and compiles it to the package's JavaScript and TypeScript declarations. The [sample workspace](../../apps/spfx-react-toolkit-test) owns the private SPFx entry points, demo UI and host configuration.

The sample consumes the compiled library. Build the library before the app; after library source changes, rebuild it and restart an active serve process before evaluating behavior. The setup guide owns the complete command sequence and configuration details.

## Choose and interpret verification

Use the verification guide to select checks for the changed behavior and its affected consumers. It distinguishes focused checks from full repository and package gates. Browser fixtures establish local rendering/interaction behavior; the SharePoint checklists cover actual host lifecycle, authentication and grants. Record those outcomes separately.

When adding or changing a public capability, keep source JSDoc, the Markdown reference and an observable sample scenario aligned. Repository instructions in [AGENTS.md](../../AGENTS.md) govern branch selection, scope, documentation/sample requirements and completion evidence.
