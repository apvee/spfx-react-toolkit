# SPFx React Toolkit

React controls, hooks, providers, typed styles, helpers and services for SharePoint Framework. Build typed forms, data grids, editable web part titles, SharePoint pickers and localized UI with runtime state isolated per SPFx instance.

![SPFx React Toolkit](https://raw.githubusercontent.com/apvee/spfx-react-toolkit/main/assets/banner.png)

[Documentation](docs/INDEX.md) · [Quick start](docs/getting-started/quick-start.md) · [API catalog](docs/api/INDEX.md) · [npm](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)

## Installation

Install in the consuming SPFx application:

```bash
npm install @apvee/spfx-react-toolkit
```

The host supplies compatible React, SPFx, PnPjs and shared Fluent/Griffel peers. Follow [installation and compatibility](docs/getting-started/installation.md) for the complete peer setup and verified baseline before building.

## Example

```tsx
import * as React from 'react';
import { useSPFxUserInfo } from '@apvee/spfx-react-toolkit';

export function Greeting(): React.ReactElement {
  const { displayName } = useSPFxUserInfo();
  return <div>Hello {displayName}!</div>;
}
```

Mount this component below `WebPartShell` in a web part, or the provider matching your actual SPFx host. The [complete quick start](docs/getting-started/quick-start.md) includes Shell integration, manifest defaults, direct provider mounting and disposal. Generic React utilities describe their individual requirements in the [API reference](docs/api/INDEX.md).

## Included APIs

| Area | Capabilities and references |
| --- | --- |
| React controls | [WebPartShell](docs/api/controls/web-part-shell.md), [text and metadata](docs/api/INDEX.md#react-controls), [forms](docs/api/controls/form.md), [DataGrid](docs/api/controls/data-grid.md), cards, feedback and Stage |
| SharePoint selections | [React pickers](docs/api/controls/pickers.md), [FileFolderPicker](docs/api/controls/file-folder-picker.md) and [native Property Pane factories](docs/api/INDEX.md#property-pane) |
| Access and host runtime | [Access trims](docs/guides/react-controls.md), [four SPFx providers](docs/api/core/providers.md), [hooks](docs/api/INDEX.md#react-and-spfx-hooks), scoped permissions and clients |
| Storage and non-React composition | [Browser/Graph/tenant/site/layered configuration](docs/guides/storage.md), [helpers and services](docs/guides/helpers-and-services.md) |
| Styling and localization | [Typed styles and useSx](docs/guides/styling.md), [native resources and strings overrides](docs/guides/localization.md) |

Each reference describes provider, shared-runtime, permission and localization requirements. Access trims control rendering; server authorization still governs data access.

## Localization setup

Add `spfx-react-toolkit configure-localization` to the consuming application's `postinstall`, preserving any existing command. Run it once after adding the setup:

```bash
npx --no-install spfx-react-toolkit configure-localization
```

Follow [localization setup](docs/guides/localization.md) for all eight resource families, custom dictionaries, `--project`, `--check` and migration from `EditableText.locale`. Native mappings remain required with complete strings overrides.

## Package imports

Root, domain and individual control imports are supported. Native Property Pane factories require dedicated imports:

```ts
import { DataGrid, createDataGrid } from '@apvee/spfx-react-toolkit/controls/data-grid';
import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
import { PropertyPaneListPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
```

See [package imports](docs/guides/package-imports.md) for resolver, shared peers, historical `/lib/` compatibility and bundle boundaries. The package includes compiled JavaScript/declarations, native localization dictionaries and the setup CLI.

## Documentation

| Task | Guide |
| --- | --- |
| Find a function, hook, provider or type | [API catalog](docs/api/INDEX.md) |
| Make HTTP, SharePoint REST or Graph calls | [HTTP clients](docs/guides/http-clients.md) |
| Select storage by lifetime and sharing | [Storage guide](docs/guides/storage.md) |
| Check SharePoint permissions or delegated scopes | [Permissions guide](docs/guides/permissions.md) |
| Compose typed CSS classes | [Styling guide](docs/guides/styling.md) |
| Work outside React | [Helpers and services](docs/guides/helpers-and-services.md) |
| Choose imports and understand bundling | [Package imports](docs/guides/package-imports.md) |

## Development

From a cloned repository's root:

```bash
npm ci
npm run build
```

The library lives in `packages/spfx-react-toolkit`; the private SPFx sample lives in `apps/spfx-react-toolkit-test` and consumes the compiled library. Rebuild the library and restart serve after library source changes. See [development setup](docs/development/setup.md) and [scoped verification commands](docs/development/verification.md).

Authenticated SharePoint behavior is verified separately using the [host checklists](docs/development/sharepoint-validation/INDEX.md). Local builds and tests do not establish tenant permissions or consent.

MIT — see [LICENSE](LICENSE) and [third-party notices](packages/spfx-react-toolkit/THIRD-PARTY-NOTICES.md). [Report an issue](https://github.com/apvee/spfx-react-toolkit/issues).
