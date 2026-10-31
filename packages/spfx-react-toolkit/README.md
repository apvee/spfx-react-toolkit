<!-- Generated from the root README.md by scripts/documentation.cjs. -->

# SPFx React Toolkit

React controls, hooks, providers, typed styles, helpers and services for SharePoint Framework. Build typed forms, data grids, editable web part titles, SharePoint pickers and localized UI with runtime state isolated per SPFx instance.

![SPFx React Toolkit](https://raw.githubusercontent.com/apvee/spfx-react-toolkit/main/assets/banner.png)

[Documentation](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INDEX.md) · [Quick start](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/getting-started/quick-start.md) · [API catalog](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/INDEX.md) · [npm](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)

## Installation

Install in the consuming SPFx application:

```bash
npm install @apvee/spfx-react-toolkit
```

The host supplies compatible React, SPFx, PnPjs and shared Fluent/Griffel peers. Follow [installation and compatibility](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/getting-started/installation.md) for the complete peer setup and verified baseline before building.

## Example

```tsx
import * as React from 'react';
import { useSPFxUserInfo } from '@apvee/spfx-react-toolkit';

export function Greeting(): React.ReactElement {
  const { displayName } = useSPFxUserInfo();
  return <div>Hello {displayName}!</div>;
}
```

Mount this component below `WebPartShell` in a web part, or the provider matching your actual SPFx host. The [complete quick start](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/getting-started/quick-start.md) includes Shell integration, manifest defaults, direct provider mounting and disposal. Generic React utilities describe their individual requirements in the [API reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/INDEX.md).

## Included APIs

| Area | Capabilities and references |
| --- | --- |
| React controls | [WebPartShell](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/web-part-shell.md), [text and metadata](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/INDEX.md#react-controls), [forms](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/form.md), [DataGrid](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/data-grid.md), cards, feedback and Stage |
| SharePoint selections | [React pickers](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/pickers.md), [FileFolderPicker](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/file-folder-picker.md) and [native Property Pane factories](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/INDEX.md#property-pane) |
| Access and host runtime | [Access trims](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/react-controls.md), [four SPFx providers](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/core/providers.md), [hooks](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/INDEX.md#react-and-spfx-hooks), scoped permissions and clients |
| Storage and non-React composition | [Browser/Graph/tenant/site/layered configuration](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/storage.md), [helpers and services](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/helpers-and-services.md) |
| Styling and localization | [Typed styles and useSx](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/styling.md), [native resources and strings overrides](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/localization.md) |

Each reference describes provider, shared-runtime, permission and localization requirements. Access trims control rendering; server authorization still governs data access.

## Localization setup

Add `spfx-react-toolkit configure-localization` to the consuming application's `postinstall`, preserving any existing command. Run it once after adding the setup:

```bash
npx --no-install spfx-react-toolkit configure-localization
```

Follow [localization setup](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/localization.md) for all eight resource families, custom dictionaries, `--project`, `--check` and migration from `EditableText.locale`. Native mappings remain required with complete strings overrides.

## Package imports

Root, domain and individual control imports are supported. Native Property Pane factories require dedicated imports:

```ts
import { DataGrid, createDataGrid } from '@apvee/spfx-react-toolkit/controls/data-grid';
import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
import { PropertyPaneListPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
```

See [package imports](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/package-imports.md) for resolver, shared peers, historical `/lib/` compatibility and bundle boundaries. The package includes compiled JavaScript/declarations, native localization dictionaries and the setup CLI.

## Documentation

| Task | Guide |
| --- | --- |
| Find a function, hook, provider or type | [API catalog](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/INDEX.md) |
| Make HTTP, SharePoint REST or Graph calls | [HTTP clients](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/http-clients.md) |
| Select storage by lifetime and sharing | [Storage guide](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/storage.md) |
| Check SharePoint permissions or delegated scopes | [Permissions guide](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/permissions.md) |
| Compose typed CSS classes | [Styling guide](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/styling.md) |
| Work outside React | [Helpers and services](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/helpers-and-services.md) |
| Choose imports and understand bundling | [Package imports](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/guides/package-imports.md) |

## Development

From a cloned repository's root:

```bash
npm ci
npm run build
```

The library lives in `packages/spfx-react-toolkit`; the private SPFx sample lives in `apps/spfx-react-toolkit-test` and consumes the compiled library. Rebuild the library and restart serve after library source changes. See [development setup](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/development/setup.md) and [scoped verification commands](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/development/verification.md).

Authenticated SharePoint behavior is verified separately using the [host checklists](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/development/sharepoint-validation/INDEX.md). Local builds and tests do not establish tenant permissions or consent.

MIT — see [LICENSE](LICENSE) and [third-party notices](https://github.com/apvee/spfx-react-toolkit/blob/main/packages/spfx-react-toolkit/THIRD-PARTY-NOTICES.md). [Report an issue](https://github.com/apvee/spfx-react-toolkit/issues).
