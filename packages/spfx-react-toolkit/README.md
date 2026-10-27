# SPFx React Toolkit

React controls, four host providers, 46 hooks, public helpers and services for SharePoint Framework. Each provider maintains runtime state for its SPFx instance. Hooks cover reusable React callbacks, typed class composition with useSx, context, properties, clients, PnPjs, themes, permissions, storage and diagnostics; helpers and services support composition outside React.

![SPFx React Toolkit](https://raw.githubusercontent.com/apvee/spfx-react-toolkit/main/assets/banner.png)

Install in an SPFx host project:

```bash
npm install @apvee/spfx-react-toolkit "@fluentui/react-components@^9.74.9" "@fluentui/react-charts@^9.3.27" "@fluentui/react-migration-v8-v9@^9.10.20" "@fluentui/react-theme@^9.2.2" "@fluentui/react-utilities@^9.26.7" "@griffel/core@^1.21.4" "@griffel/react@^1.7.8" "@fluentui/react-shared-contexts@^9.26.4"
```

The host supplies the React and SPFx runtimes. PnPjs APIs require the compatible `@pnp/core`, `@pnp/queryable` and `@pnp/sp` peers. Preserve versions compatible with your host; do not upgrade an existing SPFx toolchain just to install the toolkit.

The package declares mandatory shared RHF and Fluent compat/icon peers used by Form. The tested React 17 / Fluent 9.74.9 baseline uses the following versions; preserve the compatible Fluent/Griffel subtree in the host lockfile:

```bash
npm install react-hook-form@7.87.0 @fluentui/react-datepicker-compat@0.6.16 @fluentui/react-timepicker-compat@0.4.21 @fluentui/react-icons@2.0.312
```

See [Form dependencies and runtime identity](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/form.md#dependencies).

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

## Localized UI setup

In the consuming SPFx application's `package.json`, add `"postinstall": "spfx-react-toolkit configure-localization"` (or append it to the existing postinstall). Run once from that application after installing:

```bash
npx --no-install spfx-react-toolkit configure-localization
```

This registers the eight native resource IDs in `config/config.json`; later installs repeat setup. Controls using these resources need SPFx loader mappings even with complete `strings` overrides. A control that needs no SPFx React provider can still require native resources. See [Localization](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/LOCALIZATION.md) for monorepos, validation, culture selection and the removed `EditableText.locale` prop.

## Package imports

Root imports remain supported. Clean `/controls`, `/controls/formatted-text`, `/core`, `/hooks`, `/helpers`, `/services`, `/styles`, `/styles/useSx` and `/hooks/useStableCallback` aliases resolve to canonical implementations; historical `/lib/...` imports remain available. The `/styles` facade exports descriptors, their public types and `useSx`. For example:

```tsx
import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
import { useSx, width, typography } from '@apvee/spfx-react-toolkit/styles';
```

Use an ESM-aware bundler with package exports and sideEffects support. Shared React/Griffel/Fluent peers, theme variables and existing lifecycle limits still apply; native Node/CJS execution and all peer versions are not verified. Unrelated root imports no longer install incidental PnP registrations: direct callers must import their own PnP features, including inside invoke/batch callbacks. See [package imports and the measured production contract](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/PACKAGE-IMPORTS.md) for resolver requirements, legacy leaves, generic examples, budgets and limitations. The lazy Imports sample exercises updated captured callbacks, mixed descriptors and shared provider direction; authenticated host checks remain separate.

## Stage

[Stage](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/stage.md) presents the same Header/Body/Footer content as an explicitly selected Dialog, Drawer or anchored Popover. Small is the default size; Dialog and Drawer also support full. Controlled open requests support dirty/pending form guards, with Form.Root inside the surface and Form.Submit in its Footer. Import the single compound Stage value from the root, `/controls` or `/controls/stage`.

Stage uses existing shared Fluent/Griffel peers and inherits the host theme, direction and renderer without requiring an SPFx provider. The lazy Stage sample selects one presentation with editable size, retained state and form submission; independent-instance regressions live in the versioned fixture. Local browser execution and authenticated SharePoint/Teams/screen-reader validation are separate; host checks remain NOT EXECUTED until recorded.

## DataGrid

[DataGrid and createDataGrid](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/data-grid.md) provide sixteen native Fluent cell presentations, typed JSX or immutable descriptors, controlled sort intent and selection, and official Charts Sparkline. Consumers supply stable row IDs and own row order. Import from the root, `/controls` or `/controls/data-grid`; create typed suites outside rendering. The complete family includes Charts/D3, while unrelated imports exclude it. The mandatory Charts peer starts at `^9.3.27`.

The lazy **DataGrid** sample exercises all columns, dynamic invalid/missing values, off-page selection, resizing, loading and immutable chart updates. Authenticated SharePoint and screen-reader validation are separate checks.

## Form

For generic editable records, [Form and createForm](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/form.md) provide typed RHF state, one validation owner, 22 Fluent/native adapters, explicit reset and nested object/array sections. Their compound suite hooks do not require an SPFx provider.

## List selection

[`useSPFxPnPList`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplist) keeps its exact-title API. [`useSPFxPnPListById`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplistbyid), [`useSPFxPnPListByUrl`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplistbyurl) and [`useSPFxPnPListByPath`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/pnpjs.md#usespfxpnplistbypath) select by list GUID, decoded server-relative root URL or decoded web-relative root path. All share the existing generic item type, options, optional PnP context and CRUD/query results. They do not query on mount; invalid selectors fail lazily through existing operation failure channels (`getById` resolves `undefined` and publishes `error` for service failures; query/write actions reject). Supply the intended web's context for cross-site access; paths require an explicit client web base. The standalone [`createSPFxPnPListService`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md#createspfxpnplistservice) accepts a title string or `SPFxPnPListSelector`. Numeric item IDs remain distinct from list GUIDs.

## Site collection storage

[`useSPFxSiteKeyValueStore`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/storage.md#usespfxsitekeyvaluestore) shares one hidden `SiteKeyValueStore` in the collection root web across all subsites. The standalone [`createSPFxSiteKeyValueStoreService`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md#createspfxsitekeyvaluestoreservice) accepts `pageContext.site.absoluteUrl` explicitly. Mount, reads and removal do not provision; `save` or explicit service `ensureListReady` can provision. Effective root-web/list grants apply, and `canWrite` is advisory. Hook read failures return fallbacks with `error`; writes reject and must be caught. Existing tenant storage remains in the tenant app catalog.

## Layered configuration

[`useSPFxConfigurationStore`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/hooks/storage.md#usespfxconfigurationstore) and [`createSPFxConfigurationStoreService`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md#configuration-store-service) provide explicit tenant, site or layered configuration. Layered reads use an existing site record as a whole and fall back to tenant defaults only after confirmed absence. Construction and mount perform no reads or provisioning. Mutations require an explicit enabled scope; read and write promises reject and expose separate error/loading state in the hook.

## Package contents

The npm tarball includes `package.json`, `README.md`, `LICENSE`, `lib/index.*` and compiled `lib/core`, `lib/hooks`, `lib/helpers`, `lib/services`, `lib/utils`, `lib/styles`, `lib/controls`, `lib/property-pane` and `lib/localization` files. It also includes `THIRD-PARTY-NOTICES.md`, the `cli/spfx-react-toolkit.cjs` executable, `localization/resources.json` and the eight families' native AMD dictionaries under `lib/controls/*/loc`. JavaScript is emitted as ESNext modules, with TypeScript declarations and source/declaration maps. The main and types entries remain `lib/index.js` and `lib/index.d.ts`. Sample bundles, SPFx manifests/config, source, tests and repository docs are not shipped. Map paths reference repository source; the tarball does not embed that source.

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

Fluent and Griffel integration uses mandatory shared peer dependencies `@griffel/core` (`^1.21.4`), `@griffel/react` (`^1.7.8`), `@fluentui/react-shared-contexts` (`^9.26.4`), `@fluentui/react-migration-v8-v9` (`^9.10.20`), `@fluentui/react-theme` (`^9.2.2`), `@fluentui/react-components` (`^9.74.9`) `@fluentui/react-utilities` (`^9.26.7`) and `@fluentui/react-charts` (`^9.3.27`). The consuming project provides compatible shared packages; modern npm can install missing peers automatically. Existing compatible installations are reused. The SPFx test app declares these explicitly. Token-based styles need theme variables in scope; useSx needs no SPFx provider or new renderer. The sample uses optional `@fluentui/react-provider@9.22.8` for direction overrides that inherit the host theme. `tslib` is required by the SPFx packages that use it; this library’s ES2020 output does not import it.

See the [full styles and useSx reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/helpers/styles.md) for the typed catalog, query/state precedence, theme requirements and cache lifetime.

## Editable text

[`EditableText`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/editable-text.md) renders controlled HTML in Read mode and loads a private editor in Edit mode. Block policies constrain rendering and effective edits; changes report canonical HTML and plain text for caller persistence. Import from the root, `/controls` or `/controls/editable-text`. Configure the native resource and use `strings` for per-instance UI text; the former `locale` prop has been removed.

## Formatted text

[`FormattedText`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/formatted-text.md) renders a native text element with the existing typography catalog, restricted inline Markdown, highlight and bounded overflow. `as` defaults to `span` independently of `variant`; valid `maxLines` adds block clamp layout even on span. The control has no `inline` prop and needs no SPFx provider. Supply Fluent theme variables and shared peers, including `@fluentui/react-components` (`^9.74.9`). The library supplies unified/Remark runtime dependencies; disabling Markdown skips parsing but does not guarantee parser removal from the bundle.

```tsx
import { FormattedText } from '@apvee/spfx-react-toolkit';

<FormattedText as="h2" variant="title2">A heading</FormattedText>
<FormattedText as="p" subtle>**Secondary** text with a [link](https://example.com).</FormattedText>
<FormattedText maxLines={2} style={{ width: 240 }}>Long text in a clamped span...</FormattedText>
```

The lazy **FormattedText** sample selects one text scenario with editable typography/overflow, inherited host theme and optional direction/ref/click diagnostics. Authenticated SharePoint-host validation is separate and remains NOT EXECUTED until recorded.

## Web part shell

[`WebPartShell`](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/web-part-shell.md) wraps an actual `BaseClientSideWebPart` (`instance={this}`), supplies SPFx/Fluent providers, adapts section/application padding and edits a persisted title inline in Edit mode. Extend `IWebPartShellProperties` with application fields, compose `getWebPartShellPropertiesMetadata()` into the class getter, and use the pane's show-title toggle. The exact title properties are `title`, `titleHTML?` and `showTitle?`; defined empty HTML stays authoritative. Import from the root, `/controls` or `/controls/web-part-shell`.

The author UI uses Fluent UI 9; the private Tiptap 3.30.3 engine loads only for a visible Edit title. The locked dependency addition is 32 package records, about 7.6 MB of raw installed content (880 files), separately from emitted asset sizes. See the guide for renderer ownership, SDK subscription limits, protocol restrictions and local/authenticated validation boundaries.

## Pagination

[Pagination](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/pagination.md) supplies controlled known/unknown-total pages, cursor directions and load-more requests. The application owns rows, loading and result announcements; the control supplies accessible native Fluent navigation, localized labels, compact sizes and host theme/RTL styling. Import from the root, `/controls` or `/controls/pagination`. No new dependency or SPFx provider is required.

The lazy Pagination sample selects one controlled mode with paged results, filter/page-size/total changes, manual loading completion and optional direction diagnostics. Independent-instance regressions live in the versioned fixture. Local checks and authenticated SharePoint validation are separate; host checks remain NOT EXECUTED until recorded.

## Description List

[DescriptionList](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/description-list.md) renders read-only metadata as native term/definition pairs. Use immutable items or `DescriptionList.Item` composition for rich Fluent values. Auto stacks below 480px and aligns columns inside the nearest `container.inlineSize` ancestor; explicit layouts override this. Import the compound value from the root, `/controls` or `/controls/description-list`. Existing Fluent/Griffel peers are reused; no SPFx provider is required.

The lazy Description List sample presents one metadata list with selectable data/compound composition, layout/width/content controls and a retained child counter. Nested-container, pathological-value and independent-instance regressions live in the versioned fixture. Local verification and authenticated SharePoint/Teams/screen-reader checks are separate; host checks remain NOT EXECUTED until recorded.

## Content and preview cards

[ContentCard](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/content-card.md) composes native Fluent Card header/body/footer regions with explicit caller controls. [PreviewCard](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/preview-card.md) adds generated or custom media, src-scoped image failure, fallback and a bounded logo overlay. Both preserve native Card focus/selection and host theme/renderer; no new provider or peer is required. Import from the root, `/controls`, `/controls/content-card` or `/controls/preview-card`, respectively.

Named content accepts shorthand or fixed-tag objects/callbacks following [Control presentation slots](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/CONTROL-SLOTS.md). The control owns headings and associations; refs target the inner content containers. Strings are literal, and Markdown requires explicit FormattedText composition. Structural children and data/native peer contracts remain specific to each control. Local tests, browser execution and authenticated SharePoint/Teams/screen-reader checks provide separate evidence.

## Headless access trims

See [SharePointPermissionTrim](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/sharepoint-permission-trim.md), [SharePointGroupTrim](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/sharepoint-group-trim.md) and [EntraGroupTrim](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/controls/entra-group-trim.md) for native current-user checks, direct Loading/Fallback/ErrorFallback branches, provider requirements and advisory authorization limits. Each has an actual lazy sample panel; the Context panel also demonstrates the scoped permission hook. Graph permission requests require separate tenant approval.
