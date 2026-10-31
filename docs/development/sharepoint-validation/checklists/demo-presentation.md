# Integrated demo presentation

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated checks are **NOT EXECUTED**. Local [integrated checks](../../component-checks.md#integrated-demo-presentation-checks) simulate SPFx services/network snapshots while mounting the actual app, lazy registry panels and compiled controls.

| Check | Expected observation | Status |
| --- | --- | --- |
| Host theme and focus | Change the site/Teams theme while a note is focused; previews/popups follow tokens, drafts/counters and focus remain. No panel theme selector is present. | NOT EXECUTED |
| Navigation and disclosure | Reach all registered Fluent tabs with keyboard at 320px and native 200% zoom. Coverage summary stays visible; detailed coverage starts collapsed; diagnostics retain state and use scenario-specific initial states (AutoSkeleton, FormattedText, RelativeTime, StateMessage and Groups open; Form and Stage collapsed). | NOT EXECUTED |
| Configuration and boundaries | Fluent 9 labeled fields configure real props outside previews. Actions remain reachable; intentional overflow belongs to named bounded regions. RTL override off inherits host direction. | NOT EXECUTED |
| Form validation timing | Compare onSubmit, onBlur and onChange with a required empty Reference; errors appear at the selected event, drafts persist through mode/disclosure changes. | NOT EXECUTED |
| Isolation and requests | Two webparts retain independent values/direction and renderer context. Mounting/disclosure does not perform CRUD; explicit test actions use disposable resources. | NOT EXECUTED |

Record actual host/browser, theme/direction, screenshots, keyboard/focus, console and PASS/FAIL/NOT EXECUTED results separately from local simulation.
