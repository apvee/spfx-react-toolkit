# PropertyPaneFileFolderPicker

The native field follows the [Property Pane token bridge](pickers.md#native-pane-theme), independently of the web-part section theme. Its body-mounted Stage dialog uses the same native tokens. The existing File Picker SDK receives `default` or `dark` from that theme on each opening; an already-open Microsoft iframe is not reskinned live. Renderer, direction, focus, selections and SDK ownership remain unchanged.

The native field renders the same outlined tag surface as the React control: individually removable tags, an empty selection action and a mode-specific icon (Attach, FolderOpen or FolderDocument). It has no separate Browse/clear-all buttons. Medium/Large use the native Fluent Input border and 32px/40px minimum height; tags can wrap in a narrow pane. The focus underline remains active while the hosted picker is open. Labels, hint/error associations, disabled/error styling, host theme and reduced motion follow the shared control. This presentation does not change JSON ownership, metadata leases or hydration behavior.

Import the native factory separately from the React control:

```ts
import { PropertyPaneFileFolderPicker } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
import type { FileFolderPickerValue } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';

interface DocumentProperties {
  documents: FileFolderPickerValue[];
}

// Inside an SPFx WebPart whose properties use DocumentProperties:
// PropertyPaneFileFolderPicker<DocumentProperties>('documents', {
//   instance: this,
//   key: 'documents-picker',
//   label: 'Documents',
//   preset: 'focused',
//   value: this.properties.documents ?? [],
//   scope: { kind: 'siteCollection' },
//   multiselect: true,
//   itemTypes: 'files',
//   extensions: ['.pdf'],
// });
```

The returned value is always a JSON array of `FileFolderPickerValue`, including in single mode. `FileFolderPickerPropertyKey<P>` accepts only safe direct properties whose non-undefined type is exactly that array domain; scalar, nested, untyped and mixed-domain properties are rejected. Drive-item references are separate from the existing SharePoint list-item picker values.

The factory requires `instance`, a stable `key` and the current `value`. Optional `valueRevision` marks an authoritative external reset, including resets to an array equivalent to a previous host value. Remaining options follow [FileFolderPicker](../controls/file-folder-picker.md), excluding `value`, `defaultValue` and `onChange` from the React contract. SPFx supplies the native change callback; absent callbacks disable mutations with a diagnostic.

`initialOptions` supplies selected display models independently of saved references. Labels already available remain visible on reopening. Initially unresolved selections without labels use the shared AutoSkeleton boundary only until hydration settles. Browsing, subsequent updates and failed hydration do not repeatedly hide existing values.

The field composes the existing WebPart provider, native mount/update/dispose layer, renderer/theme/direction host and owner/document cache lease. Repeated native renders with equivalent source/value props retain the React root, pending draft and opened picker. A nonreactive pane can repeat its old property bag without discarding an accepted draft. Authoritative changes cancel pending confirmation and never emit a hydration write.

Closing the last native field releases its in-memory metadata cache and cancels its owned remote session. Reopening creates a fresh cache and uses the consumer's saved JSON values. No token storage, selection history or browser persistence is introduced. Use stable array references where practical, but equivalent recreated arrays must not cause an unnecessary remount.

The new factory uses a small direct-array value bridge because the older bridge requires `PickerValue` renderer aliases and Combobox selection events. Authentication, metadata clients, scope resolution and native hosting remain shared toolkit capabilities. Source, kind, extension, cardinality and returned-item restrictions are identical between React and Property Pane.

See [SharePoint validation](../../SHAREPOINT-VALIDATION.md) for required real-host dialog, scope, authentication and reopen checks. Local simulations do not establish tenant permissions or Microsoft's internal API traffic.

## Native member reference

`PropertyPaneFileFolderPickerProps` inherits the React source, item policy, display seeds, native Field slots and onError contracts. These native members replace React selection ownership:

| Member | Contract |
| --- | --- |
| `PropertyPaneFileFolderPickerProps.instance` | Required initialized matching WebPart instance; never serialized. |
| `PropertyPaneFileFolderPickerProps.key` | Required nonempty stable field key; never derive it from selection or time. |
| `PropertyPaneFileFolderPickerProps.value` | Required complete JSON array; use [] for an absent optional web-part property. Invalid input retains confirmed state and disables mutation. |
| `PropertyPaneFileFolderPickerProps.valueRevision` | Optional string/number reset token; a changed token adopts the supplied value even when it equals the old host property bag. |

The native field always opens the same responsive 1080 × 720 dialog without an outer visible title, close button or padding. No `presentation` option remains. `scope: { kind: 'oneDrive' }` resolves the authenticated user’s OneDrive read-only without an URL property; optional `folder` restricts that drive. Source changes still belong to the consumer and do not rewrite saved references automatically.


## Presets and opening feedback

The native factory accepts the same explicit `preset: 'focused' | 'extended'` and functional props as [React](../controls/file-folder-picker.md#presets-the-simple-configuration), with the same resolver, defaults and validation. Graph/claims handling is internal to the preset and uses existing provider/grants only. Upload and create-folder remain explicit false-by-default intents. A preset does not change accepted source or saved JSON values.

Omitting preset preserves the deprecated legacy SDK configuration. Defined legacy SDK flags cannot be mixed with a preset; choose migration explicitly. Existing web-part preferences are not silently reinterpreted, and draft/confirmed references survive presentation/preset updates.
Use `hint` for the description: omission renders none; opening never adds automatic status text or replaces the browse icon with a spinner. The animated lower border and internal busy state remain. Opening-only changes apply on the next Browse without replacing the mounted root, losing the native draft or repeating hydration. Source/policy or authoritative revision changes still cancel stale work. The final native lease disposal clears mounted metadata cache; a later pane opening hydrates with a fresh cache and preserves saved JSON references/available labels.

```tsx
PropertyPaneFileFolderPicker('documents', {
  instance: this, key: 'documents', value: this.properties.documents ?? [],
  label: 'Documents', hint: 'Project references', preset: 'focused',
  uploadEnabled: false, createFolderEnabled: false,
});
```

Optional native writes are immediate: cancelling the dialog or discarding/applying pane changes does not undo uploaded files or created folders. Persist only FileFolderPickerValue arrays; transient capabilities and callbacks never belong in the web-part property value. The sample exposes preset and explicit write intents without SDK authentication switches or tenant permission changes.

## Navigation and diagnostics

The native sample uses focused/extended with the same effective recipe as React. Focused hides native navigation; extended requests the complete source-aware navigation set. Neither changes accepted provenance. The sample clears optional initial navigation when its effective site/library/folder source changes, while preserving it for equivalent source identities, preset-only updates and unrelated OneDrive dependencies. Microsoft can display additional entries or return empty content; those outcomes need real-host verification.

Extended SharePoint browsing can lazily authenticate the exact independently verified current-user OneDrive origin. Focused does not enable auxiliary personal discovery. Existing legacy callers retain their explicit pivot-based rule. Auxiliary failures stay local to the hosted command and cannot replace a source Field error. No automatic fallback broadens scope or changes preset.

With SPFx `debug=true`, the shared implementation emits bounded internal `console.debug` authentication categories without tokens, claims, raw resource URLs or provider/native payloads. These are separate from the sanitized actionable `onError` contract; optional audience/Graph failures do not become a false configured-source failure. They introduce no public configuration member, automatic hint or spinner. Empty routes, unsupported source combinations, host policy denials, server authorization failures and unavailable browser automation targets require separate validation outcomes.

Native uploads and folder creation are immediate writes. The accepted-reference scope check is not a pre-write destination boundary, and cancelling or discarding a native draft cannot undo those operations.

### Sample Preset menu

The test web part offers only **Focused** and **Extended** in its Property Pane. Saved legacy demo configurations use Focused for both the displayed choice and the native control. Changing the preset preserves selected references, the field key and its value revision; old saved SDK fields are retained without passing them alongside a preset. The reusable Property Pane factory still supports the documented legacy contract when `preset` is omitted.

## Colorful selected-tag avatars

`PropertyPaneFileFolderPickerProps.colorfulTagAvatars?: boolean` inherits the [React FileFolderPicker contract](../controls/file-folder-picker.md#colorful-selected-tag-avatars). Default false/omitted preserves current icon avatars. True colors those existing avatars through Fluent's palette using the canonical saved identity; filenames can change without changing color. It preserves icons, sizes, the native field/session, saved references and the typed change callback. It does not style Microsoft's remote iframe or start metadata/authentication work.

```ts
PropertyPaneFileFolderPicker('documents', {
  instance: this, key: 'documents', value: this.properties.documents ?? [],
  label: 'Documents', colorfulTagAvatars: true,
});
```

The native sample has a separate **Colorful file/folder tag avatars** toggle under **File picker options**, default off.
