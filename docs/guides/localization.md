# Localization

[Documentation](../INDEX.md) · [Controls](../api/INDEX.md#react-controls)

Toolkit UI text uses native SPFx localized resources, separated by control/family. The package ships all 50 SharePoint UI cultures listed in [the locale fixture](../../tests/fixtures/sharepoint-ui-locales.json): eight resources and 400 AMD files. Canonical JSON dictionaries are validated for exact keys, nonblank text, template parameters, plural categories and calendar arrays during the library build. Translations have structural and sampled semantic review; they are not a native-speaker certification for every culture.

## Consumer installation

Add this command to the **SPFx application's** `package.json` after installing `@apvee/spfx-react-toolkit`:

```json
{
  "scripts": {
    "postinstall": "spfx-react-toolkit configure-localization"
  }
}
```

If the application already has a postinstall, append the command to its existing workflow. The library does not run a postinstall that edits parent projects. Run the command once when introducing this setup into an existing application; subsequent consumer installs rerun it. Package managers with lifecycle scripts disabled require the manual command after installation.

```bash
npx --no-install spfx-react-toolkit configure-localization
npx --no-install spfx-react-toolkit configure-localization --check
npx --no-install spfx-react-toolkit configure-localization --list
# A monorepo root can name the actual SPFx project explicitly:
npx --no-install spfx-react-toolkit configure-localization --project apps/my-spfx-app
```

The command reads the installed package's inventory, enumerates all resources and adds missing entries to `config/config.json` → `localizedResources`. It resolves the package used by the selected consumer, including hoisted installations, and writes paths relative to that consumer. Existing mappings and unrelated configuration are preserved. Repeating setup without changes leaves the file byte-for-byte unchanged. Configuration must be strict JSON; an absent or invalid project/configuration fails with an actionable error.

`--check` does not write: it fails when a required ID has no nonblank mapping. It checks registration, not the contents or existence of a consumer's custom translations. `--list` prints the inventory. To repair a stale generated path after moving dependencies, remove the affected toolkit mapping and rerun setup. Custom mappings are intentionally never replaced automatically.

## Resource families

| Native resource ID | Packaged path after `node_modules/@apvee/spfx-react-toolkit/` |
| --- | --- |
| `ApveeSpfxReactToolkitPaginationStrings` | `lib/controls/pagination/loc/{locale}.js` |
| `ApveeSpfxReactToolkitDataGridStrings` | `lib/controls/data-grid/loc/{locale}.js` |
| `ApveeSpfxReactToolkitEditableTextStrings` | `lib/controls/editable-text/loc/{locale}.js` |
| `ApveeSpfxReactToolkitWebPartShellStrings` | `lib/controls/web-part-shell/loc/{locale}.js` |
| `ApveeSpfxReactToolkitPickersStrings` | `lib/controls/pickers/loc/{locale}.js` |
| `ApveeSpfxReactToolkitFileFolderPickerStrings` | `lib/controls/file-folder-picker/loc/{locale}.js` |
| `ApveeSpfxReactToolkitFormStrings` | `lib/controls/form/loc/{locale}.js` |
| `ApveeSpfxReactToolkitStageStrings` | `lib/controls/stage/loc/{locale}.js` |

Each owning feature imports its own native module. Setup registers all IDs so future imports work without further configuration; SPFx determines the resources needed by the compiled component. Registration alone does not cause a React runtime to load all translations. `en-us.js` is the packaged default. Use the supported SPFx build/loader localization behavior for culture selection and fallback; changing a React prop does not reconfigure the native loader.

A consumer can map an ID to its own compatible AMD files, for example `lib/loc/toolkit-pagination/{locale}.js`. Preserve the native resource schema, template parameters and plural forms. The published inventory is discoverable metadata; resource schemas are private implementation contracts. Prefer the typed React dictionaries below for ordinary customization.

## Per-instance dictionaries

Controls accept typed `strings` overrides. Omitted fields inherit the SPFx resource. A complete dictionary supplies another language explicitly, including callback formatters where the public type requires them. Dictionaries affect UI text; they do not translate stored content, option labels or backend data. No global localization provider is required.

```tsx
import * as React from 'react';
import { EditableText, Pagination, Stage, createForm } from '@apvee/spfx-react-toolkit';

<Pagination
  pagination={{ type: 'pages', pageIndex: 0, pageSize: 10, totalItems: 25 }}
  strings={{ nextPage: 'Next report page' }}
/>;

<Stage.Header title="Report" strings={{ closeLabel: 'Dismiss report' }} />;
// Stage.Header must be mounted inside its Stage surface.

<EditableText value="<h2>Report</h2>" mode="read"
  strings={{ formattingControls: 'Report formatting' }} />;

const ReportForm = createForm<{ title: string }>();
function ReportEditor(): React.ReactElement {
  const methods = ReportForm.useForm({
    defaultValues: { title: '' },
    strings: {
      unexpectedFailure: 'The report could not be saved.',
      fileInput: { browseLabel: 'Attach report' },
    },
  });
  return <ReportForm.Root methods={methods} onSubmit={() => { throw new Error('Service unavailable'); }}>
    <ReportForm.Input name="title" label="Title" />
    <ReportForm.Error />
    <ReportForm.Submit>Save</ReportForm.Submit>
  </ReportForm.Root>;
}
```

Form dictionaries belong to the `useForm` instance. A Root with external `methods` inherits that dictionary; an initializing Root can receive its own `strings`. Dictionaries become visible after commit. Async unexpected failures use the latest committed text when the failure is created. Already stored errors keep their original message; changing dictionaries does not reset fields or rerun validation. Rule, schema and resolver-provided domain messages remain caller-owned.

Direct label/content props take precedence over family dictionaries. Empty-value behavior follows each API's established contract: Pagination trims labels and falls back for blank text; Stage, editor and Shell preserve explicit empty text; nested picker/Form dictionaries use `undefined` for inheritance. Slot suppression and direct leaf calendar/array/file labels retain their documented precedence. See the [control guides](../api/INDEX.md#react-controls) for exact fields and callbacks.

## Formatting and external authorities

UI resource culture is independent of regional formatting. DataGrid root/column `locale`, RelativeTime `locale`, timezone selection, parsers, value formatters and `Intl` behavior remain in place. HTML `lang`/`dir`, native browser file dialogs and Microsoft hosted File Picker language keep their own authorities. FileFolderPicker `strings` customizes toolkit-owned outer text and safe error messages; it does not translate Microsoft's iframe or raw transport diagnostics.

## Migration and other hosts

`EditableText.locale` has been removed. Replace it with a partial or complete `EditableTextStringsOverride`, or omit `strings` to use SPFx defaults. Shell title naming uses `WebPartShellStrings`; formatting text is supplied through the editor dictionary. Updating dictionaries preserves live editor state and does not publish content. This is a breaking migration from the earlier locale-based editor API; use the dictionary types exported by the installed package.

The production controls now require the native SPFx resource IDs at build/runtime. Typed overrides still merge with those defaults and do not make the current package independent of SPFx. A future plain-React distribution would need an explicit adapter/build boundary for those IDs. Repository standalone browser tests use actual emitted AMD data through a documented webpack adapter; that test adapter is not a production loader or evidence of authenticated SharePoint behavior.

See [SharePoint validation](../development/sharepoint-validation/INDEX.md) for the remaining real-host checks.

## Verify native UI cultures

For an Italian debug document, append `locale=it-it` to the actual hosted-workbench URL and inspect the native family resource requests. Compare Context's UI locale with its regional locale; they can differ. Keep explicit DataGrid/RelativeTime formatting locales independent of resource labels, and test simultaneous instance dictionaries without changing persisted content.

Follow the [hosted bootstrap procedure](../development/setup.md#hosted-workbench-bootstrap-and-debug-culture) to check actual registration and rendering. Use the [native localization checklist](../development/sharepoint-validation/checklists/localization.md) for culture, cold-start and unsupported-culture scenarios, recording fresh observations separately from local fixtures.

## CLI option contract

| Input | Behavior |
| --- | --- |
| No option | Configures the consumer in the current working directory, adding only missing toolkit resource IDs. |
| `--project <directory>` | Selects the actual consumer directory; a relative path resolves from the current working directory. May appear once. |
| `--check` | Checks strict configuration and nonblank registrations without writing. Does not inspect custom AMD contents or repair generated paths. |
| `--list` | Prints resource IDs, packaged path patterns and supported cultures from the CLI's inventory; requires no consumer configuration or installed consumer package. |

`--check` and `--list` are mutually exclusive. Duplicate flags, missing project directories, unknown options/commands, invalid inventory or invalid consumer mappings fail with a nonzero exit status before configuration is written. For configuration/check mode the executing CLI installation must match the package resolved for the chosen consumer; use that consumer's installed executable. Setup preserves existing custom/unknown mappings, indentation, line endings, trailing-newline convention and property order. It does not build the library, grant permissions or configure a standalone non-SPFx loader.

The native resource generator/formatters are private implementation modules rather than public helper APIs. The published `localization/resources.json` inventory is packaged metadata for discovery/setup. Consumers should use typed control dictionary contracts for overrides instead of importing internal resource formatter functions.

## Locale and time-zone metadata

`getSPFxLocaleInfo(pageContext)` and `useSPFxLocaleInfo()` expose regional `locale`, native `uiLocale`, `isRtl`, raw host `timeZone`, and optional representative `ianaTimeZone`. The IANA value derives only from a verified SharePoint numeric group; unknown/excluded IDs stay undefined. It is independent of resource-loader culture and is not a physical-city or user-profile time zone. See the [complete locale helper contract](../api/helpers/page-context.md#getspfxlocaleinfo) for mapping provenance, exclusions, custom SharePoint Server rules and an explicit missing-zone policy.

```ts
const info = getSPFxLocaleInfo(pageContext);
if (info.ianaTimeZone !== undefined) {
  const formatter = new Intl.DateTimeFormat(info.locale, { timeZone: info.ianaTimeZone });
  console.log(formatter.format(new Date()));
}
```

Using undefined as an Intl `timeZone` selects the runtime's zone; make that fallback policy explicit in the application. A locale chooses formatting conventions and native UI text authorities choose translations; neither alone establishes the site's regional time-zone rules.
