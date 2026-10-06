# SPFx React Toolkit

React providers, 40 hooks, public helpers and services for SharePoint Framework. Each provider maintains runtime state for its SPFx instance. Hooks cover context, properties, clients, PnPjs, themes, permissions, storage and diagnostics; helpers and services support composition outside React.

![SPFx React Toolkit](https://raw.githubusercontent.com/apvee/spfx-react-toolkit/main/assets/banner.png)

Install in an SPFx host project:

```bash
npm install @apvee/spfx-react-toolkit "@fluentui/react-migration-v8-v9@^9.9.12" "@fluentui/react-theme@^9.2.0"
```

The host supplies the React and SPFx runtimes. PnPjs APIs require the compatible `@pnp/core`, `@pnp/queryable` and `@pnp/sp` peers. Preserve versions compatible with your host; do not upgrade an existing SPFx toolchain just to install the toolkit.

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFxWebPartProvider, useSPFxUserInfo } from '@apvee/spfx-react-toolkit';

const Greeting: React.FC = () => {
  const { displayName } = useSPFxUserInfo();
  return <div>Hello {displayName}!</div>;
};

export default class GreetingWebPart extends BaseClientSideWebPart<{}> {
  public render(): void {
    ReactDom.render(
      React.createElement(SPFxWebPartProvider, { instance: this }, React.createElement(Greeting)),
      this.domElement
    );
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

Choose `SPFxWebPartProvider`, `SPFxApplicationCustomizerProvider`, `SPFxFieldCustomizerProvider` or `SPFxListViewCommandSetProvider` to match the actual SPFx host. The sample includes real WebPart and Application Customizer entry points. Field Customizer and Command Set provider coverage currently checks exports; mounting those providers needs their corresponding hosts.

## Repository development

This repository uses npm workspaces:

- `packages/spfx-react-toolkit`: publishable library; `tsc` emits ESNext modules, declarations and maps to `lib`.
- `apps/spfx-react-toolkit-test`: private SPFx sample; SPFx 1.21.1 Gulp build performs Sass, lint and webpack bundling against the library package.
- `docs`, `scripts`, `tests`: public documentation, verification and regression tests.
- `.docs`: internal maintenance records and local agent planning material.

Run from the cloned repository root:

```bash
npm ci
npm run build:library
npm run build:app
npm run verify
npm run verify:package
```

`npm test` runs the Node behavioral suite; it is not a Gulp test task. `npm run build` builds library then app. To debug locally:

```bash
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve
```

After library source changes, rebuild with `npm run build:library` and restart the app's serve process. The app consumes the package's compiled `lib` entry point. See [Development](./docs/DEVELOPMENT.md) for the full command map and [SharePoint validation](./docs/SHAREPOINT-VALIDATION.md) for real-host checks and prerequisites.

The layout migration preserves the package import name and public API. Source moved from root `src` to the library workspace; SPFx sample source/config moved to the app workspace. Root build output is replaced by workspace output. Consumers continue to import `@apvee/spfx-react-toolkit`.

## Requirements

| Requirement | Repository verification baseline |
|-------------|----------------------------------|
| Node.js | `>=22.14.0 <23.0.0` |
| SPFx build/runtime packages | `1.21.1` |
| React / ReactDOM | `17.0.1` (package peers: `17.x`) |
| TypeScript | `5.3.3` |
| PnPjs | `4.17.0` (package peers: `^4.0.0`) |

The package declares SPFx peers `>=1.18.0 <2.0.0`; that range is a compatibility declaration, not evidence of testing every version. Repository checks use SPFx 1.21.1. Authenticated tenant validation is a separate step.

## Documentation

- [Introduction and quick start](./docs/INTRODUCTION.md)
- [API reference: 4 providers, 40 hooks, helpers and services](./docs/INDEX.md)
- [Helpers API](./docs/api/helpers/INDEX.md)
- [Services API](./docs/api/services/INDEX.md)
- [NPM package](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)
- [Issues](https://github.com/apvee/spfx-react-toolkit/issues)

MIT — see [LICENSE](./LICENSE).

Fluent integration uses mandatory peer dependencies `@fluentui/react-migration-v8-v9` (`^9.9.12`) and `@fluentui/react-theme` (`^9.2.0`). The consuming project provides compatible shared packages; modern npm can install missing peers automatically. Existing compatible installations are reused. The SPFx test app declares both explicitly. `tslib` is required by the SPFx packages that use it; this library’s ES2020 output does not import it.
