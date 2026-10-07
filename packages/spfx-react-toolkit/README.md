# SPFx React Toolkit

React providers, 46 hooks, public helpers and services for SharePoint Framework. Each provider maintains runtime state for its SPFx instance. Hooks cover reusable React callbacks, typed class composition with useSx, context, properties, clients, PnPjs, themes, permissions, storage and diagnostics; helpers and services support composition outside React.

![SPFx React Toolkit](https://raw.githubusercontent.com/apvee/spfx-react-toolkit/main/assets/banner.png)

Install in an SPFx host project:

```bash
npm install @apvee/spfx-react-toolkit "@fluentui/react-migration-v8-v9@^9.9.12" "@fluentui/react-theme@^9.2.0" "@fluentui/react-utilities@^9.25.1" "@griffel/core@^1.19.2" "@griffel/react@^1.5.30" "@fluentui/react-shared-contexts@^9.25.2"
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

## Package imports

Root imports remain supported. Clean `/core`, `/hooks`, `/helpers`, `/services`, `/styles`, `/styles/useSx` and `/hooks/useStableCallback` aliases resolve to canonical implementations; historical `/lib/...` imports remain available. The `/styles` facade exports descriptors, their public types and `useSx`. For example:

```tsx
import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
import { useSx, width, typography } from '@apvee/spfx-react-toolkit/styles';
```

Use an ESM-aware bundler with package exports and sideEffects support. Shared React/Griffel/Fluent peers, theme variables and existing lifecycle limits still apply; native Node/CJS execution and all peer versions are not verified. Unrelated root imports no longer install incidental PnP registrations: direct callers must import their own PnP features, including inside invoke/batch callbacks. See [package imports and the measured production contract](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/PACKAGE-IMPORTS.md) for resolver requirements, legacy leaves, generic examples, budgets and limitations. The lazy Imports sample exercises updated captured callbacks, mixed descriptors and shared provider direction; authenticated host checks remain separate.

## List selection

[`useSPFxPnPList`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplist) keeps its exact-title API. [`useSPFxPnPListById`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplistbyid), [`useSPFxPnPListByUrl`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplistbyurl) and [`useSPFxPnPListByPath`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplistbypath) select by list GUID, decoded server-relative root URL or decoded web-relative root path. All share the existing generic item type, options, optional PnP context and CRUD/query results. They do not query on mount; invalid selectors fail lazily through existing operation failure channels (`getById` resolves `undefined` and publishes `error` for service failures; query/write actions reject). Supply the intended web's context for cross-site access; paths require an explicit client web base. The standalone [`createSPFxPnPListService`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md#createspfxpnplistservice) accepts a title string or `SPFxPnPListSelector`. Numeric item IDs remain distinct from list GUIDs.

## Site collection storage

[`useSPFxSiteKeyValueStore`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/storage.md#usespfxsitekeyvaluestore) shares one hidden `SiteKeyValueStore` in the collection root web across all subsites. The standalone [`createSPFxSiteKeyValueStoreService`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md#createspfxsitekeyvaluestoreservice) accepts `pageContext.site.absoluteUrl` explicitly. Mount, reads and removal do not provision; `save` or explicit service `ensureListReady` can provision. Effective root-web/list grants apply, and `canWrite` is advisory. Hook read failures return fallbacks with `error`; writes reject and must be caught. Existing tenant storage remains in the tenant app catalog.

## Package contents

The npm tarball includes `package.json`, `README.md`, `LICENSE`, `lib/index.*` and compiled `lib/core`, `lib/hooks`, `lib/helpers`, `lib/services`, `lib/utils` and `lib/styles` files. JavaScript is emitted as ESNext modules, with TypeScript declarations and source/declaration maps. The main and types entries remain `lib/index.js` and `lib/index.d.ts`. Sample bundles, SPFx manifests/config, source, tests and repository docs are not shipped. Map paths reference repository source; the tarball does not embed that source.

The repository moved library source to `packages/spfx-react-toolkit/src` and the SPFx sample to `apps/spfx-react-toolkit-test`. The import name, entry points and public API are unchanged by this layout migration.

## Requirements

| Requirement | Repository verification baseline |
|-------------|----------------------------------|
| Node.js | `>=22.14.0 <23.0.0` |
| SPFx build/runtime packages | `1.21.1` |
| React / ReactDOM | `17.0.1` (package peers: `17.x`) |
| TypeScript | `5.3.3` |
| PnPjs | `4.17.0` (package peers: `^4.0.0`) |

The package declares SPFx peers `>=1.18.0 <2.0.0`; that range is a compatibility declaration, not evidence of testing every version. Repository checks use SPFx 1.21.1. Authenticated tenant validation is a separate step.

## Development from a clone

Clone [the repository](https://github.com/apvee/spfx-react-toolkit) and run commands from its root:

```bash
npm ci
npm run build:library
npm run build:app
npm run verify
npm run verify:package
npm run pack:library
```

`pack:library` rebuilds the library and writes an npm tarball to root `artifacts`. It does not publish. Local SPFx debugging uses:

```bash
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve --workspace @apvee/spfx-react-toolkit-test
```

Rebuild the library and restart serve after changing library source. These development scripts belong to the cloned repository, not an installed npm package.

- [Introduction](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INTRODUCTION.md)
- [API reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INDEX.md)
- [Development guide](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/DEVELOPMENT.md)
- [SharePoint validation guide](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/SHAREPOINT-VALIDATION.md)
- [Helpers API](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/helpers/INDEX.md)
- [Services API](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md)

MIT — see [LICENSE](./LICENSE).

Fluent and Griffel integration uses mandatory shared peer dependencies `@griffel/core` (`^1.19.2`), `@griffel/react` (`^1.5.30`), `@fluentui/react-shared-contexts` (`^9.25.2`), `@fluentui/react-migration-v8-v9` (`^9.9.12`), `@fluentui/react-theme` (`^9.2.0`) and `@fluentui/react-utilities` (`^9.25.1`). The consuming project provides compatible shared packages; modern npm can install missing peers automatically. Existing compatible installations are reused. The SPFx test app declares these explicitly. Token-based styles need theme variables in scope; useSx needs no SPFx provider or new renderer. The sample uses optional `@fluentui/react-provider@9.22.8` for its theme controls. `tslib` is required by the SPFx packages that use it; this library’s ES2020 output does not import it.

See the [full styles and useSx reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/helpers/styles.md) for the typed catalog, query/state precedence, theme requirements and cache lifetime.
