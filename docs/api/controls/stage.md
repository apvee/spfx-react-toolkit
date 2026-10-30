# Stage

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > Stage

`Stage` presents one content structure as a modal Dialog, modal Drawer or anchored nonmodal Popover. Choose `presentation` explicitly; viewport changes never switch it. Compose `Stage.Header`, `Stage.Body` and optional `Stage.Footer` in declaration order. Stage needs the shared Fluent theme/Griffel renderer, but no SPFx provider or new dependency.

```tsx
import * as React from 'react';
import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';

<Stage presentation="dialog" trigger={<button type="button">Edit profile</button>}>
  <Stage.Header title="Profile" description="Edit your details" closeLabel="Close" />
  <Stage.Body><input aria-label="Display name" /></Stage.Body>
  <Stage.Footer><button type="button">Help</button></Stage.Footer>
</Stage>
```

Import from the package root, `/controls` or `/controls/stage`; all share the canonical implementation. `Stage` is the only value export, with `Stage.Header`, `Stage.Body` and `Stage.Footer` compound members. There are no flat part exports or public Trigger. The fourteen public types, including the `StageStrings` localization dictionary, are documented below; state-ownership and surface-attribute helper types remain internal to the public contract.

## StageProps

Use one Header and one Body, with an optional Footer. Children may be inside an application wrapper; Stage neither collects nor reorders them. Parts require an owning Stage. The root and each part forward an `HTMLDivElement` ref. Native surface attributes/events, `className`, consumer `style`, `dir`, `lang`, `hidden` and `id` are forwarded. Stage owns `role`, `aria-modal`, `aria-labelledby` and `aria-describedby`; root `title`, `dangerouslySetInnerHTML` and duplicate root header presentation slots are unsupported.

| Member | Type | Behavior/default |
| --- | --- | --- |
| `StageProps.presentation` | `StagePresentation` | Required discriminant; dialog, drawer or popover. |
| `StageProps.size` | `StageSize` for dialog/drawer; excludes full for popover | small. |
| `StageProps.trigger` | `React.ReactElement` | Optional for dialog/drawer, required for popover; native button or Fluent Button. |
| `StageProps.drawerPosition` | `StageDrawerPosition` | Drawer only; end. Unavailable for dialog/popover. |
| `StageProps.positioning` | `StagePopoverPositioning` | Popover only; below/start. Unavailable for dialog/drawer. |
| `StageProps.withArrow` | `boolean` | Popover only; false. Shows the native Fluent anchor arrow and updates while open. Unavailable for dialog/drawer. |
| `StageProps.open` | `boolean` | Authoritative controlled state; mutually exclusive with defaultOpen. |
| `StageProps.defaultOpen` | `boolean` | Initial uncontrolled state; false. |
| `StageProps.children` | `React.ReactNode` | Required caller-owned content. |
| `StageProps.onOpenChange` | `(event: StageOpenChangeEvent, data: StageOpenChangeData) => void` | One notification for an accepted interaction request. |
| `StageProps.dismissMode` | `StageDismissMode` | standard. |
| `StageProps.unmountOnClose` | `boolean` | true; false retains compatible content after first open. |
| `StageProps.mountNode` | `DialogSurfaceProps['mountNode']` | Optional Fluent portal destination; omit to use the provider/document defaults. |

Select controlled or uncontrolled ownership once per mount; changing ownership throws. Invalid presentation/size and full Popover are rejected even for untyped callers. Use `undefined` for absent optional values.

## StageHeaderProps

`Stage.Header` renders a native div around a neutral Fluent Card/CardHeader. The inner Card isolates ancestor Card context, including selectable Cards, without adding a second visible Card. Native attributes, ref and handlers belong to the outer div; they do not enable the inner Card's hover/pressed interaction styling. Native div attributes are supported except `title`, `children` and `dangerouslySetInnerHTML`.

| Member | Type | Behavior/default |
| --- | --- | --- |
| `StageHeaderProps.title` | `span` content slot | Required nonempty phrasing content; heading wrapper uses subtitle2/primary. Strings have a one-line ellipsis limit. |
| `StageHeaderProps.description` | `span` content slot | Optional supporting phrasing content; caption1/subtle. Strings have a two-line ellipsis limit. |
| `StageHeaderProps.icon` | `span` content slot | Decorative 24px slot, hidden from assistive technology. |
| `StageHeaderProps.action` | `div` content slot | Optional caller action, before the owned Close button; no automatic dismissal. |
| `StageHeaderProps.headingLevel` | `StageHeadingLevel` | 2; semantic level independent of visual typography. |
| `StageHeaderProps.showCloseButton` | `boolean` | true; hiding X does not change dismissal policy. |
| `StageHeaderProps.closeLabel` | `string` | Overrides `strings.closeLabel` and the host-selected native label. Explicit empty/whitespace values are preserved. |

Strings are literal: `**Profile**` remains `**Profile**`, with no Markdown parsing. In Dialog, Drawer and Popover, a direct title string or a title props object's string `children` occupies at most one line with ellipsis; equivalent description strings occupy at most two lines with ellipsis. Truncation is visual: the complete text remains in the DOM, and the owned heading/description IDs and accessible associations are preserved. Text block margins and title-description gap are zero. JSX and render callback output remain caller-owned presentation, retain their props, refs, handlers and explicit styling, and inherit wrapper presets when they supply no override; Stage does not apply the string line limits to them. Do not nest headings or block markup in the title.

Use the singular `action` prop for the header's application action. This replaces `actions` without a compatibility alias; migrate existing `Stage.Header actions={...}` usages to `action={...}`. Action and Close stay on the same row and are vertically centered with the title/description block in Dialog, Drawer and Popover, including narrow layouts and RTL. Close remains separately owned at the logical end, with reserved space. The action column is bounded to leave room for text; caller controls must allow shrinking (for example, a Button with `sx(minWidth.zero)` and a label span with `sx(minWidth.zero, text.truncate)`), retaining full accessible label text. Prefer one compact visible action or an application-owned menu. If Close is hidden on a modal surface, provide a visible way to close it. Footer actions and other controls' plural `actions` props keep their existing contracts.

## StageBodyProps and StageFooterProps

`Stage.Body` is the growing scroll region. `Stage.Footer` holds arbitrary actions/content; its absence leaves no empty footer gap. Header and Footer do not shrink. Both forward native div attributes/events, refs and consumer classes/style, excluding `dangerouslySetInnerHTML`.

| Member | Type | Behavior |
| --- | --- | --- |
| `StageBodyProps.children` | `React.ReactNode` | Optional body content within available space. |
| `StageFooterProps.children` | `React.ReactNode` | Optional free JSX; actions do not close automatically. |

The surface is a flex column. Application wrappers must supply their own compatible flex layout. At viewport heights up to 359px the whole surface can scroll, keeping Header/Footer reachable when they would crowd Body.

For an action footer, use `sx(flex.row, flex.wrap, alignItems.center, justifyContent.end, gap.small)`: actions align to the logical end (right in LTR, left in RTL), wrap when needed and use Fluent S spacing tokens (8px in the standard themes), matching the gap of Fluent DialogActions. Footer content remains caller-owned. A form wrapping the parts should use `gap.medium` to keep the same M spacing (12px in the standard themes) as the surface.

## StagePresentation, StageSize and StageDrawerPosition

`StagePresentation` is `'dialog' | 'drawer' | 'popover'`. `StageSize` is `'small' | 'medium' | 'large' | 'full'`, narrowed by StageProps for Popover. `StageDrawerPosition` is `'start' | 'end'`; end is the default logical edge and follows provider RTL direction.

| Size | Dialog | Drawer | Popover |
| --- | --- | --- | --- |
| small (default) | 640px | 320px | 320px |
| medium | 800px | 592px | 480px |
| large | 1120px | 940px | 640px |
| full | Viewport width and height minus two outer gaps | Full viewport width and available height | Unsupported |

Popover uses preferred widths of 320/480/640px, reduced by native Fluent auto-size when viewport space is smaller. The default remains `small`, now 320px. This replaces the previous Popover scale of 240/320/480px; existing small and medium Popovers become wider, and large increases to 640px. Dialog and Drawer presets retain their existing widths. Padding remains 12px per side for Popover. Select each size in the test web part and resize the viewport to observe the available-space constraint in LTR and RTL.

Dialog's equal outer gaps are `clamp(8px, 1vw, 16px)` on all sides. Measure full/narrow geometry against the usable CSS viewport: Fluent's native modal body-scroll lock applies `scrollbar-gutter: stable`, which can reserve a classic scrollbar gutter. Consequently CSS `100vw` can differ from `window.innerWidth`; do not treat that difference as a Stage sizing failure. The responsive gap still follows CSS `1vw`, bounded from 8px to 16px. Full fills both dimensions inside that gap; other Dialog sizes follow content height up to the available maximum. Drawer fills available height and reaches the selected viewport edge, including full size. Popover retains native anchoring, flip/shift and auto-size within viewport limits. Dynamic viewport height is used when supported with vh fallback. Modal surface padding is 24px; Popover padding is 12px. Header/Body/Footer add no duplicate padding.

## StageHeadingLevel and StagePopoverPositioning

Use `withArrow` with Popover to show Fluent's native arrow. Toggle the boolean while open; content and form state stay mounted. Native positioning keeps the arrow pointed toward the trigger, including flip/shift and RTL. The native surface keeps the arrow outside its border while an inner wrapper scrolls content on constrained viewports.

```tsx
<Stage presentation="popover" withArrow={showArrow} trigger={<button type="button">Details</button>}>
  <Stage.Header title="Details" />
  <Stage.Body>Contextual content</Stage.Body>
</Stage>
```

`StageHeadingLevel` accepts 1 through 6, default 2. The wrapper heading supplies Stage's accessible title association; a description is associated only while content is actually present.

| Member | Type | Behavior/default |
| --- | --- | --- |
| `StagePopoverPositioning.position` | `'above' \| 'below' \| 'before' \| 'after'` | below; preferred logical side. |
| `StagePopoverPositioning.align` | `'start' \| 'center' \| 'end'` | start; preferred alignment. |

Dialog/Drawer use native modal dialog semantics and focus management. Popover uses the native nonmodal group surface without a focus trap. Supply a named native `<button>` or Fluent `Button` trigger; arbitrary elements/custom wrappers are unsupported. Stage preserves trigger ref, disabled state, name and event handlers. It owns `aria-expanded` and `aria-controls`; Popover omits `aria-haspopup` because its surface is a group. Consumer activation cancellation prevents opening. Native focus initial placement/restoration is Fluent-owned; removed triggers are not a usable restore destination.

## StageDismissMode, StageOpenChangeReason, StageOpenChangeEvent and StageOpenChangeData

`StageDismissMode` is `'standard' | 'explicit'`. Standard allows Escape and applicable backdrop/outside/focus-outside interactions. Explicit filters those requests before notification, permitting only trigger and close-button requests; the Popover trigger still toggles closed. Header and Footer actions never dismiss automatically.

`StageOpenChangeReason` is `'triggerClick' | 'closeButton' | 'escapeKeyDown' | 'backdropClick' | 'outsideClick' | 'focusOutside'`. `StageOpenChangeEvent` is a React MouseEvent, KeyboardEvent or FocusEvent of HTMLElement, or a native MouseEvent, TouchEvent, FocusEvent or KeyboardEvent.

| Member | Type | Behavior |
| --- | --- | --- |
| `StageOpenChangeData.open` | `boolean` | Requested next state, which can differ from controlled open. |
| `StageOpenChangeData.reason` | `StageOpenChangeReason` | Normalized interaction source. |

Uncontrolled requests update state and notify once; prop changes do not notify. Controlled consumers accept by updating `open`, or refuse dirty/pending dismissal by retaining it. Calling `event.preventDefault()` in onOpenChange is not a veto. Fluent can already prevent native Popover Escape before delivering its callback; Stage still accepts that request. An Escape consumed first by a nested Menu/Combobox/DatePicker must dismiss the child first. The installed legacy DatePicker compat input stops propagation and prevents Escape even when its calendar is already closed. Repeated Escape while focus remains on that closed input therefore does not close Stage. Use X or an application footer close action, or move focus to another Stage control before pressing Escape. Stage respects child consumption and does not rewrite Form or third-party keyboard handlers. Nested Fluent portals preserve native virtual-parent relationships for outside interactions.

Presentation is latched for each open session. A requested mode change applies next session, retaining the active adapter's compatible trigger and placement meanwhile. Size can change live; a pending full Dialog/Drawer leaves an active Popover at its last compatible size. Refused controlled closes keep the session intact. Do not depend on child-state retention across adapter changes.

With `unmountOnClose=true`, native presence/motion removes content on close. With false, content first mounts when opened, then remains hidden, aria-hidden and inert while closed; mounted effects remain active. Compatible size, theme and action rerenders preserve child state. Adapter changes can reset retained content.

## Form composition

Put Form.Root **inside** the portaled surface so the inputs and footer Submit belong to one DOM form. React context alone from an outer Form.Root does not establish DOM form ownership. Submit through the suite's Submit/Root gate rather than calling methods.handleSubmit in the Footer.

```tsx
import * as React from 'react';
import { createForm, Stage, useSx, alignItems, flex, flexItem, justifyContent, minHeight, width, gap } from '@apvee/spfx-react-toolkit';

const Editor = createForm<{ name: string }>();

export function ProfileEditor(): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const [discardRequested, setDiscardRequested] = React.useState(false);
  const methods = Editor.useForm({ defaultValues: { name: '' } });
  // Read during render to subscribe to RHF's proxied formState before callbacks use it.
  const { isDirty, isSubmitting } = methods.formState;
  const sx = useSx();
  return <Stage presentation="dialog" open={open}
    trigger={<button type="button">Edit profile</button>}
    onOpenChange={(_, data) => {
      if (!data.open && isSubmitting) return;
      if (!data.open && isDirty) { setDiscardRequested(true); return; }
      setOpen(data.open);
    }}>
    <Editor.Root methods={methods}
      className={sx(flex.column, flexItem.grow, minHeight.zero, width.full, gap.medium)}
      onSubmit={values => {
        methods.reset(values);
        setDiscardRequested(false);
        setOpen(false);
      }}>
      <Stage.Header title="Profile" closeLabel="Close" />
      <Stage.Body>
        <Editor.Input name="name" label="Display name" />
        {discardRequested && <div role="status">
          Discard changes?
          <button type="button" onClick={() => {
            methods.reset(); setDiscardRequested(false); setOpen(false);
          }}>Discard</button>
          <button type="button" onClick={() => setDiscardRequested(false)}>Keep editing</button>
        </div>}
      </Stage.Body>
      <Stage.Footer className={sx(flex.row, flex.wrap, alignItems.center, justifyContent.end, gap.small)}>
        <Editor.Submit appearance="primary">Save</Editor.Submit>
      </Stage.Footer>
    </Editor.Root>
  </Stage>;
}
```

For persistence, await the application save in Root's onSubmit before reset/close. Keep pending requests controlled and use an inline discard confirmation rather than nesting a modal. See [Form](form.md) for validation, async submission and reset semantics. Stage has no RHF dependency of its own.

## Dependencies, portals and validation

Stage uses the existing shared Fluent components/icons and Griffel peers. The repository baseline is React/ReactDOM 17.0.1, TypeScript 5.3.3, Fluent 9.74.9 and SPFx 1.21.1, with Node `>=22.14.0 <23.0.0`; it introduces no dependency or peer upgrade. Inherit theme, direction, targetDocument and renderer from the enclosing Fluent provider, including an iframe's owner document. Stage creates no global renderer/provider and adds no arbitrary z-index.

Omitting mountNode retains Fluent's themed portal/virtual-parent behavior. An explicit mountNode is a caller-owned escape hatch: use the correct owner document and verify theme variables, direction, renderer styles, focus and nested portal relationships at that destination. Do not automatically move content to document.body.

The lazy Stage tab shows one configurable two-field editor in Dialog, Drawer or Popover. Presentation, size, title and description drive the same active surface; header action and Close stay aligned, while save/dismiss/retention guards remain observable. Configuration, open trigger, compact status and advanced diagnostics share one outline card. Modal-only manual completion remains inside the surface to preserve focus ownership. Nested popup, long-content and independent/RTL/native-ancestor regressions remain in the [versioned Stage fixture](../../../tests/fixtures/demo-control-isolation/stage-regressions.tsx). The product and conformance browser results, including any console gate failure, must be reported separately from authenticated host observations.

## Named content slots

Titles, descriptions and icons accept React content or a native `span` props object; `action` accepts React content or a native `div` props object. Use `action` on `Stage.Header`. Props objects support attributes, classes, consumer styles, refs and `children: (Component, props) => JSX` callbacks. The callback receives the fixed inner tag and composed props during rendering; inspection never executes it. Forward those props to preserve its attributes and ref.

```tsx
<Stage.Header
  title={{ as: 'span', className: 'custom-title', children: 'Preferences' }}
  description={{ children: 'Literal **supporting text**' }}
  action={<button type="button">Help</button>}
/>
```

The heading, generated IDs, surface associations and Close button remain owned by Stage. A slot ref points to its inner container. Strings/JSX shorthand retain their existing DOM and literal formatting; props objects explicitly request an inner container even with empty children. Booleans and absent values suppress content; zero remains visible. Tags cannot replace the owned heading, surface or Close button; raw HTML is unsupported. Title, description and icon content must remain valid phrasing content; put application controls in `action`. The existing Fluent theme/host renderer is still required.

A description props object with empty or whitespace-only static children adds no accessible surface-description association unless the object supplies aria-label/aria-labelledby. Render callbacks request that association without probing their output. Shorthand, including fragments, preserves the existing presence policy: empty text/fragments are absent, while whitespace text is present. Object/callback icon containers fill the owned shell; their direct SVG children retain the 24px header sizing, even with larger authored SVG dimensions. Callback output and custom components remain caller-owned and opaque.

## Native UI localization

SPFx selects the Stage native resource from the host UI culture after the host registers its toolkit resource ID through the [localization configuration CLI](../../guides/localization.md). The native resource is a static module dependency: outside the SPFx loader, supply an explicit resource-module adapter even when passing `strings` or `closeLabel`. No toolkit provider is required, but overrides do not remove that module dependency.

`Stage.Header` accepts `strings?: Partial<StageStrings>`; the only owned label is `StageStrings.closeLabel`, the close button accessible name. Omitted/undefined dictionary values use that native label. An explicit direct `closeLabel` prop takes precedence over the dictionary, including empty or whitespace values. Dictionary strings also preserve explicit empty/whitespace values. These overrides are independent per Header.

```tsx
<Stage presentation="dialog">
  <Stage.Header title="Profile" strings={{ closeLabel: 'Dismiss profile' }} />
  <Stage.Body>Profile details</Stage.Body>
</Stage>
```

Title, description, action content and close-request callbacks keep their existing contracts.

| API | Type | Behavior |
| --- | --- | --- |
| `StageHeaderProps.strings` | `Partial<StageStrings>` | Optional per-Header dictionary over the native baseline. |
| `StageStrings.closeLabel` | `string` | Close-button accessible name; the direct prop wins. |
