# DataGrid

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated SharePoint checks and screen-reader observations below are **NOT EXECUTED**. Build the library before the app, restart serve, then open the lazy **DataGrid** panel. Source/compiled browser and standalone tarball results establish local behavior separately. See [DataGrid](../../../api/controls/data-grid.md).

| Host check | Expected result | Status |
| --- | --- | --- |
| Sixteen columns and controlled rows | Typed, dynamic and descriptor grids render; sort headers report intent before Apply sort changes supplied order. | NOT EXECUTED |
| Selection and row identity | Select-all/deselect-all retains off-page IDs; supply/remove rows without pruning selection; callbacks preserve original IDs. | NOT EXECUTED |
| Native keyboard and actions | Arrow navigation and group Enter/Escape remain native. Links/actions/custom buttons do not activate rows; activation grid handles unprevented click/Enter. | NOT EXECUTED |
| Width, resize and zoom | Header/body widths align before/after resizing. Long action labels remain usable; 320px region contains horizontal overflow at 200% zoom. | NOT EXECUTED |
| Charts and overflow | Immutable points/color updates change actual Charts. Tags/people popovers recover hidden content; theme, RTL and forced colors remain usable. | NOT EXECUTED |
| Loading, errors and lifecycle | Five rectangular skeleton rows appear without data; refresh retains existing rows; consumer Retry clears the error; all-hidden status detaches the native ref. | NOT EXECUTED |
| Screen reader | Headers, selection labels, boolean icon alternatives and chart summary announce meaning once; decorative charts add no focus stop. | NOT EXECUTED |
| Host instance isolation | Two web parts retain independent controlled state/theme/renderer context and dispose native widget subscriptions on removal. | NOT EXECUTED |
