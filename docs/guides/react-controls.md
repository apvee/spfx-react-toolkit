# Choosing and composing React controls

[Documentation](../INDEX.md) > React controls

The toolkit's controls compose the host's React 17, Fluent UI 9 and Griffel runtime. Start with the control that owns the interaction you need, then keep application data, requests and persistence in the consumer. The [API reference](../api/INDEX.md#react-controls) is the complete export catalog; the references below explain behavior and configuration.

## Choose a control

| Need | Control and contract |
| --- | --- |
| Render styled text and restricted inline Markdown | [FormattedText](../api/controls/formatted-text.md); choose semantic `as` separately from visual typography. |
| Show a localized elapsed instant | [RelativeTime](../api/controls/relative-time.md); pass an explicit zoned instant, with a controlled reference or live clock. |
| Keep content mounted behind geometry-derived placeholders | [AutoSkeleton](../api/controls/auto-skeleton.md); requires measurable bounded HTML and native inert for replacement. |
| Explain an empty/error/success section with actions | [StateMessage](../api/controls/state-message.md); supply application copy and explicit callback/link actions. |
| Summarize content with a heading, body and controls | [ContentCard](../api/controls/content-card.md); native Card interaction and independent header/footer actions. |
| Add a generated image or custom preview to a card | [PreviewCard](../api/controls/preview-card.md); explicit preview precedence, stable failure area and bounded logo. |
| Display labeled properties of one record | [DescriptionList](../api/controls/description-list.md); data or compound Item composition with native dl/dt/dd semantics. |
| Group preferences or expandable settings | [SettingsGroup](../api/controls/settings-group.md) and [AccordionGroup](../api/controls/accordion-group.md); caller-owned state, with optional Form field-row naming. |
| Edit validated application values | [Form](../api/controls/form.md); typed RHF suites, 22 adapters, custom bindings, objects and record arrays. |
| Compare records in a table | [DataGrid](../api/controls/data-grid.md); sixteen typed cell presentations, controlled sort/selection and native grid navigation. |
| Request another result page or batch | [Pagination](../api/controls/pagination.md); pages, cursors and load-more requests with consumer-owned fetching/results. |
| Open a Dialog, Drawer or anchored Popover | [Stage](../api/controls/stage.md); explicit presentation, compound header/body/footer and controlled dismissal requests. |
| Edit one formatted HTML region | [EditableText](../api/controls/editable-text.md); controlled HTML/text callbacks, finite block policy and a lazy editor. |
| Supply SPFx providers and a persisted canvas title | [WebPartShell](../api/controls/web-part-shell.md); requires an actual initialized web-part instance. |
| Select SharePoint entities or principals | [SharePoint pickers](../api/controls/pickers.md); seven controls with display models separate from JSON references. |
| Select existing SharePoint/OneDrive Business files and folders | [FileFolderPicker](../api/controls/file-folder-picker.md); Microsoft hosted v8 dialog, verified source restrictions and explicit write opt-ins. |
| Mount content only after a current-user access check | [SharePointPermissionTrim](../api/controls/sharepoint-permission-trim.md), [SharePointGroupTrim](../api/controls/sharepoint-group-trim.md) and [EntraGroupTrim](../api/controls/entra-group-trim.md); separate loading, denial and error branches. |

## Host dependencies and imports

Controls are available from the package root, `/controls` and the documented family entry, such as `@apvee/spfx-react-toolkit/controls/content-card`. These aliases share one implementation. Compound parts belong to their family: use `Stage.Header`, `SettingsGroup.Item`, `AccordionGroup.Item` and `DescriptionList.Item`. Create `createForm` and `createDataGrid` suites outside rendering; each suite owns its methods/column markers and descriptors. See [package imports](package-imports.md) for published paths, declaration dependencies and bundle boundaries.

Ordinary presentation, Form, DataGrid, Stage and Pagination controls do not require an SPFx context provider. They still require their declared shared peers and the host Fluent theme/direction/Griffel renderer. Form additionally uses its RHF/compat peers; DataGrid includes its Charts peer. An individual family import does not guarantee that only one member of a compound family is retained by the bundler.

The seven SharePoint pickers, FileFolderPicker and all three access trims require the matching [SPFx provider](../api/INDEX.md#providers-and-core-types). `WebPartShell` installs its own web-part and Fluent providers for an actual web-part instance. Do not wrap the shell in a second toolkit provider merely to make its children use toolkit hooks. Native [Property Pane factories](../api/INDEX.md#property-pane) own separate provider/render roots and must be imported through their documented Property Pane paths.

Native UI localization is independent of the SPFx provider requirement. Form, DataGrid, Stage, Pagination, WebPartShell, EditableText, SharePoint pickers and FileFolderPicker statically load their own native resource modules. Register them in an SPFx solution with the [localization setup](localization.md). A non-SPFx host needs an explicit resource-module adapter even when passing a complete `strings` dictionary; overrides alone do not remove the static resource dependency. RelativeTime uses browser Intl locale/time-zone data instead.

## Ownership and composition

| Interaction | The control owns | The consumer owns |
| --- | --- | --- |
| Form submission | RHF validation, pending gate, field state, reset/draft coordination and error presentation | Initial values, validation rules/schema/resolver, service calls and record changes |
| DataGrid sorting and selection | Native navigation, request callbacks and cell presentation | Item order, selected-ID Set, hidden/sort state, filtering, fetching and pagination |
| Pagination | Available-action UI and a discriminated request | Page/cursor state, request cancellation, pending locks, rows and result announcements |
| Stage | Surface semantics, compound associations, native focus integration and dismissal requests | Explicit presentation, acceptance of controlled requests, save/discard policy and persistence |
| EditableText | Safe one-region Read/Edit UI and effective user-change notifications | Saved HTML, controlled echoes and external persistence |
| Picker selection | Identity/source verification, UI selection, hydration and documented caches | Controlled keys/JSON arrays, display seeds and eventual storage |
| Access trim | Current-user reads and mounting only the active structural branch | Protected operation authorization on the server and caller-owned branch content |

Place Form.Root inside a portaled Stage surface when its footer submits the form: React context across a portal does not make an outer native form own the portaled inputs. Await application persistence inside Root's `onSubmit` before resetting values or closing the surface. A controlled Stage rejects a close request by keeping `open`; preventing the callback event is not a Stage veto. See the complete [Stage/Form example](../api/controls/stage.md#form-composition).

DataGrid preserves supplied item order. Pair it with Pagination only after the consumer applies sorting/filtering and supplies the corresponding rows. Page changes are intent, so set `loading` while an asynchronous request is pending and commit the new result/state deliberately. Selection can retain off-page IDs; do not silently reconstruct it from the currently visible rows.

## Presentation and absence

Named presentation regions use the [control slot standard](control-slots.md): shorthand retains its DOM, while an object requests the documented inner tag and may supply native attributes, classes, refs or a Fluent children render callback. Semantic headings, triggers, IDs and accessible associations stay with the control. Keep noninteractive phrasing inside labels/headings/triggers and put explicit controls in the owning action region. Strings are literal unless the owning reference explicitly documents Markdown; compose FormattedText for parsing elsewhere.

Use undefined for new absence/default state. Each reference defines its own presence/default policy: an empty object can request a region, boolean/empty shorthand can suppress it, and some legacy nullish inputs retain defaults. A defined PreviewCard preview overrides image even when empty. Form's built-in date/time/file/SpinButton domains retain the null required by their native/compat APIs. Do not infer one universal empty-value rule across controls.

Use `useSx` descriptors or the host's Griffel classes for authored layout and dynamic styles. Existing consumer `style` props remain escape hatches. Slot refs target inner containers; a control's public root ref continues to target its documented root. Forward the callback-supplied props so composed refs, classes and attributes survive.

## Local examples and real-host validation

The private sample's lazy control tabs exercise actual components: Cards, Description List, Pagination, DataGrid, Form, Stage, SettingsGroup/AccordionGroup, StateMessage, AutoSkeleton, RelativeTime, EditableText, FormattedText, pickers and all three trims. Native Property Pane examples exercise array persistence and cascading entity sources. Versioned conformance fixtures preserve broader state, native-comparison and isolation cases separately from those product examples.

Build the library before the app because the sample imports compiled `lib`; restart a running serve process after library source changes. [Development verification](../development/verification.md) and [browser checks](../development/browser-checks.md) distinguish static/Node evidence from rendered checks. Local tests, mock transports and builds do not establish authenticated SharePoint, Graph consent, native Apply/save/reopen, keyboard/screen-reader behavior or Microsoft's hosted picker traffic. Record those outcomes through the [SharePoint validation guides](../development/sharepoint-validation/INDEX.md).
