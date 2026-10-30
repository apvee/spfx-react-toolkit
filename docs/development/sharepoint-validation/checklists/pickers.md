# Picker host acceptance

[Validation index](../INDEX.md) · [Setup](../setup.md)

The sample **SharePoint pickers** panel and **Picker references (JSON)** native group require real host verification beyond local mocked transports:

- Query all seven domains with empty/failed/partial responses; verify scoping in a subweb, another web/list and hub association.
- Change consumer options, labels, presentation and callbacks while querying; preserve input/selections and avoid equivalent REST/client initialization.
- Exercise reactive and nonreactive Apply; explicit valueRevision reset, pane close/reopen, stale dispose, theme/RTL/iframe document isolation.
- With already-authorized delegated permissions, query People.Read `/me/people` and User.Read `/me/memberOf`; test denied/limited membership independently and the five-suggestion total. `/me/people` remains in maintenance mode intentionally.
- Resolve users in available/Entra/SharePoint/both modes; ensureUser=false missing registration blocks confirmation. With an explicitly authorized test user, confirm ensureUser=true only writes on selection/ID action. No grants or provisioning are implicit in this checklist.
- Verify library files/folders, custom key uniqueness, field type icons, calculated/Note/non-text searches and native paging against tenant constraints.

## Complete catalogs, contextual choices and query feedback

- Verify Fields/Views across all metadata pages and case-insensitive local searches beyond 30 entries, with unchanged REST counts while typing. ListPicker retains its configurable initial cap.
- Confirm current collection and actual associated hub choices, no-association/self-hub exclusion, search deduplication, selected contextual rows and cached reopen.
- Distinguish pending, empty and no-additional-results feedback while preserving current choices. Verify partial retrieval, retry, scope changes and stale responses.
- Verify native light/dark pane tokens separately from a contrasting web-part section, including body-mounted menus, retained IDs and selections, and no theme-triggered queries. Record high contrast and live host theme coverage separately.
- Compare default and colorful selected-tag avatars across React/native pickers and Form.TagPicker. Changing appearance must retain selected values, dirty state and request counts; verify photo/icon precedence and RTL/narrow layouts.

## Dropdown presentation and native lifecycle

Authenticated SharePoint/Teams/Property Pane and screen-reader checks for the new Dropdown presentation are **NOT EXECUTED**. Local compiled Fluent/React fixture checks use simulated SDK/REST responses and do not establish tenant behavior.

- Exercise all seven React and native fields with historical `presentation="combobox"`, single/multiple and medium/large controls. Verify the persistent selection button, popup search, fixed empty invitation, search placeholder overrides and selection count.
- Search while selected, clear only the query, reactivate the current single row, replace/clear single selection, and remove multi selections from Selezionati. Compare TagPicker: selected tags must not be repeated in suggestion, search or contextual rows.
- Verify arrows/Enter, editable Home/End/spaces, IME, Escape focus restoration and Tab/Shift+Tab without an implicit commit. Internal trigger/search/clear/action focus transitions must not run blur validation; outside exit must validate once.
- Delay and refuse selection validation: pending focus remains stable; refusal retains query/popup without a change; acceptance closes single only. Dismiss and reopen during pending work, checking outside focus and the newer popup session.
- Test light/dark native pane theme, RTL, narrow viewport, long labels/metadata, visible clear/chevron/media, disabled/required/searchable=false, limit, partial feedback, retry and loadMore. Body portals must inherit their own document/theme.
- Reopen the real nonreactive pane, Apply an accepted draft, discard another, and check authoritative reset and initial hydration without rewriting saved JSON. Record actual tenant/account, browser, permissions and assistive technology separately.
