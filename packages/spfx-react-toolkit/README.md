# SPFx React Toolkit

React controls, hooks, providers, helpers and services for SharePoint Framework (SPFx), with Fluent UI 9. Build web parts with editable titles, typed forms and data grids, SharePoint pickers and localized UI, while keeping runtime state isolated per SPFx instance.

![SPFx React Toolkit](https://raw.githubusercontent.com/apvee/spfx-react-toolkit/main/assets/banner.png)

[Included APIs](#included-apis) · [Installation](#installation) · [Localization](#configure-localization) · [Quick start](#quick-start) · [Imports](#package-imports) · [Documentation](#documentation)

## Included APIs

| Area | Capabilities | Guides |
| --- | --- | --- |
| Web part shell | `WebPartShell`: SPFx and Fluent providers, editable title, Card appearance and host spacing | [WebPartShell](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/web-part-shell.md) |
| Text and metadata | `FormattedText`, `EditableText`, `RelativeTime`, `DescriptionList` | [Controls reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/INDEX.md) |
| Loading and feedback | `AutoSkeleton`, `StateMessage` | [Skeletons](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/auto-skeleton.md), [states](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/state-message.md) |
| Forms and settings | `Form` / `createForm`, `SettingsGroup`, `AccordionGroup` | [Forms](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/form.md), [settings](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/settings-group.md), [disclosure](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/accordion-group.md) |
| Tables and navigation | `DataGrid` / `createDataGrid`, `Pagination` | [DataGrid](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/data-grid.md), [Pagination](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/pagination.md) |
| Cards and surfaces | `ContentCard`, `PreviewCard`, `Stage` with Dialog, Drawer or Popover presentation | [ContentCard](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/content-card.md), [PreviewCard](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/preview-card.md), [Stage](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/stage.md) |
| SharePoint selection | Sites, lists, fields, content types, views, items, principals and existing files/folders | [React pickers](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/pickers.md), [file/folder picker](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/file-folder-picker.md) |
| Property Pane | Native fields for the SharePoint pickers and file/folder picker | [Picker fields](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/property-pane/pickers.md), [file/folder field](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/property-pane/file-folder-picker.md) |
| Access checks | `SharePointPermissionTrim`, `SharePointGroupTrim`, `EntraGroupTrim` | [Access trims](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/INDEX.md#headless-access-trims) |
| React and SPFx runtime | Four host providers; hooks for properties, clients, PnPjs, themes, permissions, storage and diagnostics; `useStableCallback` and typed styles with `useSx` | [Providers](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/core/providers.md), [hooks](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/INDEX.md), [styles](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/helpers/styles.md) |
| Helpers and services | Non-React composition, list operations and tenant/site/layered configuration stores | [Helpers](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/helpers/INDEX.md), [services](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md) |

Access trims control what is rendered; server-side permissions still govern access to data. Each control's guide documents its provider, theme and localization requirements.

## Installation

Install in the **consuming SPFx application**:

```bash
npm install @apvee/spfx-react-toolkit
```

The host supplies compatible React, ReactDOM and SPFx runtimes. Keep those versions aligned with the host rather than upgrading its toolchain just to install the toolkit. The package also declares `@pnp/core`, `@pnp/queryable` and `@pnp/sp` v4 peers for its PnPjs APIs.

Add missing shared UI peers using compatible versions. All packages below are mandatory package peers, including Charts and the Form adapters; reuse compatible installations already present in the application:

```bash
npm install \
  "@fluentui/react-components@^9.74.9" \
  "@fluentui/react-charts@^9.3.27" \
  "@fluentui/react-migration-v8-v9@^9.10.20" \
  "@fluentui/react-theme@^9.2.2" \
  "@fluentui/react-utilities@^9.26.7" \
  "@fluentui/react-shared-contexts@^9.26.4" \
  "@griffel/core@^1.21.4" \
  "@griffel/react@^1.7.8" \
  react-hook-form@7.87.0 \
  @fluentui/react-datepicker-compat@0.6.16 \
  @fluentui/react-timepicker-compat@0.4.21 \
  @fluentui/react-icons@2.0.312
```

Modern npm can install missing peers automatically. Review the resolved versions and preserve shared React/Fluent/Griffel identities; see [Form dependencies](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/form.md#dependencies) for the tested adapter baseline.

### Compatibility

| Dependency | Declared contract | Repository verification baseline |
| --- | --- | --- |
| SPFx runtime peers, including Property Pane | `>=1.18.0 <2.0.0` | `1.21.1` |
| React / ReactDOM | `17.x` | `17.0.1` |
| PnPjs | `^4.0.0` | `4.17.0` |
| Node.js | `>=22.14.0 <23.0.0` | `22.20.0` |
| TypeScript | Host toolchain | `5.3.3` |

The declared SPFx range is broader than the verified baseline; it does not imply testing every SPFx release. Theme-based controls and styles need Fluent theme variables and the shared host renderer.

## Configure localization

Add the localization command to the **consumer application's** `package.json`:

```json
{
  "scripts": {
    "postinstall": "spfx-react-toolkit configure-localization"
  }
}
```

Append it if the application already has a postinstall. Run it once after adding this setup to an existing project; later consumer installs repeat registration:

```bash
npx --no-install spfx-react-toolkit configure-localization
```

The CLI enumerates the installed resource inventory and adds missing mappings to `config/config.json` → `localizedResources`. The package supplies eight control/family resources covering 50 SharePoint UI cultures. Registration preserves existing mappings and unrelated configuration; the library does not edit parent projects through its own postinstall.

SPFx selects the UI resource culture. Typed per-instance `strings` overrides customize toolkit labels independently of the formatting `locale` props used by `DataGrid`, `RelativeTime` and `Intl`. Native resource mappings are still required with complete overrides, including for controls that do not need an SPFx React provider.

When upgrading, replace the removed `EditableText.locale` with `strings` overrides where needed; formatting locales on other controls remain available. See [Localization](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/LOCALIZATION.md) for custom dictionaries, resource IDs, monorepo `--project` setup and the read-only `--check` command.

## Quick start

Use `WebPartShell` in a web part `.tsx` file. It supplies the SPFx and Fluent providers, so its children can use the toolkit hooks:

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart, type IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { PropertyPaneToggle, type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import {
  WebPartShell,
  getWebPartShellPropertiesMetadata,
  useSPFxProperties,
  useSPFxUserInfo,
  type IWebPartShellProperties,
} from '@apvee/spfx-react-toolkit';

interface IExampleWebPartProperties extends IWebPartShellProperties {
  description: string;
}

function Content(): React.ReactElement {
  const { displayName } = useSPFxUserInfo();
  const { properties } = useSPFxProperties<IExampleWebPartProperties>();
  return <p>Hello {displayName}! {properties?.description}</p>;
}

export default class ExampleWebPart extends BaseClientSideWebPart<IExampleWebPartProperties> {
  public render(): void {
    ReactDom.render(
      <WebPartShell instance={this} appearance="outline" size="medium" titlePosition="outside">
        <Content />
      </WebPartShell>,
      this.domElement
    );
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      ...getWebPartShellPropertiesMetadata(),
      description: { isSearchablePlainText: true },
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ groups: [{ groupName: 'Title', groupFields: [
      PropertyPaneToggle('showTitle', { label: 'Show title', checked: this.properties.showTitle ?? true }),
    ] }] }] };
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

Initialize these defaults in the web part manifest's `preconfiguredEntries[].properties`. Leave `titleHTML` omitted until an author edits the title:

```json
{
  "title": "My web part",
  "showTitle": true,
  "description": "Built with SPFx React Toolkit."
}
```

### Shell presentation

| Option | Default | Behavior |
| --- | --- | --- |
| `appearance` | `subtle` | No Card wrapper with `subtle`; `filled`, `filled-alternative` and `outline` render a Fluent Card. |
| `size` | `medium` | Native Card size: `small`, `medium` or `large`; ignored with `subtle`. |
| `titlePosition` | `outside` | `inside` places the complete title/action header inside the Card; ignored with `subtle`. |
| Persisted `showTitle` property | `true` when omitted | `false` hides the entire title/editor/action header and its gap. |

`surface` / `padding` control outer host spacing; Card size controls inner spacing. `appearance`, `size` and `titlePosition` are React props: exposing them in the consumer's Property Pane is optional. See [WebPartShell](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/web-part-shell.md) for title persistence, host detection, actions and state retention.

For direct provider composition, choose `SPFxWebPartProvider`, `SPFxApplicationCustomizerProvider`, `SPFxFieldCustomizerProvider` or `SPFxListViewCommandSetProvider` for the matching SPFx host. See [Providers](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/core/providers.md) and [Introduction](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INTRODUCTION.md). Generic React controls and utility hooks document their individual provider requirements.

## Package imports

Root imports are supported, together with public domain and individual control entry points:

```ts
import { DataGrid, createDataGrid } from '@apvee/spfx-react-toolkit/controls/data-grid';
import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
import { useSx, width, typography } from '@apvee/spfx-react-toolkit/styles';
import { PropertyPaneListPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
```

Property Pane factories use dedicated `/property-pane/pickers` and `/property-pane/file-folder-picker` imports; they are not root value exports. Preserved historical `/lib/...` paths remain available. Use an ESM-aware bundler with package exports and `sideEffects` support for tree shaking. Direct PnPjs callers must import their own PnP feature modules. See [Package imports](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/PACKAGE-IMPORTS.md) for resolver requirements, aliases and measured bundle boundaries.

The npm package contains compiled ESNext JavaScript, TypeScript declarations/maps, native AMD localization dictionaries and the setup CLI. Library source, sample bundles and repository documentation remain in the repository.

## Documentation

- [Introduction and provider examples](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INTRODUCTION.md)
- [Complete API reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INDEX.md)
- [Controls and public types](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/INDEX.md)
- [Localization and migration](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/LOCALIZATION.md)
- [Package imports and tree shaking](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/PACKAGE-IMPORTS.md)
- [Development guide](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/DEVELOPMENT.md)
- [SharePoint validation guide](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/SHAREPOINT-VALIDATION.md)
- [npm package](https://www.npmjs.com/package/@apvee/spfx-react-toolkit) · [Issues](https://github.com/apvee/spfx-react-toolkit/issues)

## Develop from a clone

The repository uses npm workspaces: `packages/spfx-react-toolkit` owns the library and `apps/spfx-react-toolkit-test` owns the private SPFx sample. Run these commands from the cloned repository root:

```bash
npm ci
npm run build
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve
```

The sample consumes the library's compiled `lib` entry point. After library source changes, rebuild the library and restart serve. Use targeted checks for the affected behavior; full verification and packaging commands are documented in [Development](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/DEVELOPMENT.md#verification-scope). `Verify monorepo` runs only when manually requested from GitHub Actions; see [workflow setup](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/DEVELOPMENT.md#github-actions) for availability and execution details.

Local builds and tests are separate from authenticated SharePoint-host validation. Follow [SharePoint validation](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/SHAREPOINT-VALIDATION.md) to exercise the sample in a real host.

MIT — see [LICENSE](./LICENSE) and [third-party notices](./THIRD-PARTY-NOTICES.md).
