# Authenticated SharePoint validation

[Documentation home](../../INDEX.md) · [Development](../INDEX.md) · [Setup](setup.md) · [Record results](#record-results)

This section describes repeatable manual integration checks in a real, authenticated SharePoint host. Use it to observe provider lifecycle, remote requests, effective permissions, persistence and styling in the environment where the toolkit will run.

Local tests, builds and simulated transport results establish their own checks; they do not establish tenant authentication, permission grants or actual host behavior. Each checklist becomes evidence only when a dated execution records what was observed in a specific environment and account.

## Prepare a validation run

1. Follow the [host setup](setup.md) for the test tenant, account, sample build, local serve configuration and authenticated entry points.
2. Choose the scenarios affected by the change using the table below. Check each scenario's prerequisites and prepare disposable test data where needed.
3. Record expected and actual outcomes, evidence and cleanup for each scenario. Keep unavailable scenarios visible as NOT EXECUTED or BLOCKED.

## Choose a scenario

| Scenario | Checklist |
| --- | --- |
| Provider synchronization, two instances and cleanup | [Host and lifecycle](checklists/host-lifecycle.md) |
| Concurrent calls, lists and search | [Async, list and search](checklists/async-lists-search.md) |
| List title, ID, URL and path selectors | [List selector modes](checklists/list-selectors.md) |
| Graph consent, OneDrive and tenant storage | [Graph and storage permissions](checklists/graph-storage-permissions.md) |
| Collection-root storage and schema/concurrency boundaries | [Site collection store](checklists/site-key-value-store.md) |
| Themes, classes, queries and scrollbars | [Styles and useSx](checklists/styling.md) |
| Aliases, callbacks and renderer/provider identity | [Imports](checklists/imports.md) |

## Record results

For every row record **PASS**, **FAIL**, **NOT EXECUTED** or **BLOCKED**, timestamp, host/account role, expected/actual result, steps, redacted console/network evidence and cleanup. Keep local automated results separate from authenticated tenant observations. Remove only disposable keys/items/files created for the run. Any deployment, tenant permission approval or publication requires its own authorized workflow.

| Status | Meaning for the recorded scenario |
| --- | --- |
| PASS | The observed behavior matched the expected result in the recorded environment. |
| FAIL | The observed behavior differed from the expected result; include reproduction details and evidence. |
| NOT EXECUTED | The scenario has not been run. A documented procedure or local simulation is not an execution. |
| BLOCKED | A specific prerequisite or external condition prevented the run; record what is missing. |

Save dated reports and evidence under `.docs/maintenance/`. Each result applies to its recorded host, version, account and scenario. These public procedures remain usable without access to local reports. Only disposable, authorized test data belongs in a validation run.
