<!-- Generated from the root README.md by scripts/documentation.cjs. -->

# SPFx React Toolkit

React providers, hooks, typed styles, helpers and services for SharePoint Framework. Access context, properties, clients, PnPjs, themes, permissions and storage with runtime state isolated per SPFx instance.

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

Mount this component below the provider matching your actual SPFx host. The [complete quick start](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/getting-started/quick-start.md) includes a web part, provider mounting and disposal. Generic React utilities describe their individual requirements in the [API reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/INDEX.md).

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

MIT — see [LICENSE](LICENSE). [Report an issue](https://github.com/apvee/spfx-react-toolkit/issues).
