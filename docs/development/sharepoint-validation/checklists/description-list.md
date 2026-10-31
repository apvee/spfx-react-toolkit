# Description List

[Validation index](../INDEX.md) · [Setup](../setup.md)

Status: **NOT EXECUTED** in an authenticated SharePoint/Teams host. Build library before app, restart any active serve process and open the lazy **Description List** panel. See the [API reference](../../../api/controls/description-list.md). Local tests and browser fixtures are separate evidence.

| Scenario | Expected observation | Status |
| --- | --- | --- |
| Semantics and composition | One dl with dt/dd pairs for items or JSX; rich Badge/link/paragraph slots retain markup. Zero is visible and blank/undefined retain empty definitions. Strings remain literal. | NOT EXECUTED |
| Container layout | Auto stacks at 320/479px and aligns at 480/639/640/900px; explicit layouts override. Nested 320px container remains vertical. Long labels/values wrap. Verify real workbench, page sections and Teams iframe widths. | NOT EXECUTED |
| State, focus and refs | Increment A/B, reorder stable IDs and change layout/width; state and focused child persist within the selected API. Inspect root dl and compound grouping div refs; native click events bubble. | NOT EXECUTED |
| Hidden and independence | Root and first compound pair hidden remove their rendered content. Two instances inherit the host theme and retain independent layout/counter state and explicit direction overrides. | NOT EXECUTED |
| Theme, RTL and accessibility | Verify light/dark, RTL logical order, forced colors and zoom. Screen reader conveys each visible term with its definition; links/buttons remain keyboard-operable without added focus stops. | NOT EXECUTED |
