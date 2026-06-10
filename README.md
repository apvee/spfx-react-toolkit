# SPFx React Toolkit

> A comprehensive React runtime, hooks, helpers, and services library for SharePoint Framework (SPFx). Simplifies SPFx development with instance-scoped state isolation, ergonomic hooks, and reusable non-React composition APIs across WebParts, Extensions, and Command Sets.

![SPFx React Toolkit](./assets/banner.png)

---

## Overview

**SPFx React Toolkit** is a production-ready library that simplifies SharePoint Framework development by providing unified React providers, strongly-typed hooks, and reusable public helpers/services.

Built on a provider-scoped runtime store, it delivers per-instance state isolation, automatic synchronization, and APIs that work both inside React hooks and in non-hook composition code.

### Key Benefits

| Benefit | Description |
|---------|-------------|
| 💪 **Type-Safe** | Full TypeScript support across hooks, helpers, and services |
| ⚡ **Optimized** | Provider-scoped runtime state with per-instance isolation |
| 🔄 **Auto-Sync** | Bidirectional synchronization between React and SPFx |
| 🎨 **Universal** | Works with WebParts, Application Customizers, Field Customizers, and Command Sets |
| 📦 **Modular** | Public hooks for React usage, plus helpers/services for non-hook composition |

### Features

- ✅ **40 React Hooks** — Comprehensive API surface for SPFx runtime data, clients, token providers, and API permission prechecks
- ✅ **Public Helpers** — Pure utilities for storage keys, property pane helpers, environment checks, logging, and context extraction
- ✅ **Public Services** — Reusable service factories for storage, SPFx context data, properties, theme, PnPjs, API permission prechecks, and provider composition
- ✅ **Instance Isolation** — State scoped per SPFx instance (multi-instance support)
- ✅ **PnPjs Integration** — Optional hooks for PnPjs v4 with type-safe filters
- ✅ **Cross-Platform** — Teams, SharePoint, and Local Workbench support

---

## Quick Start

### Installation

Install the package in your SPFx project:

```bash
npm install @apvee/spfx-react-toolkit
```

If you use the PnPjs hooks or services, install the peer PnPjs packages as well:

```bash
npm install @pnp/core @pnp/queryable @pnp/sp
```

Then wrap the SPFx entry point with the provider that matches the component type you are building:

- `SPFxWebPartProvider` for WebParts
- `SPFxApplicationCustomizerProvider` for Application Customizers
- `SPFxFieldCustomizerProvider` for Field Customizers
- `SPFxListViewCommandSetProvider` for ListView Command Sets

### Basic Usage

```typescript
// In your WebPart
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';

public render(): void {
  const element = (
    <SPFxWebPartProvider instance={this}>
      <MyComponent />
    </SPFxWebPartProvider>
  );
  ReactDom.render(element, this.domElement);
}

// In your component
import { useSPFxProperties, useSPFxUserInfo } from '@apvee/spfx-react-toolkit';

const MyComponent: React.FC = () => {
  const { properties } = useSPFxProperties<IMyProps>();
  const { displayName } = useSPFxUserInfo();

  return <div>Hello {displayName}!</div>;
};
```

### Helpers and Services

Use public helpers and services when the logic must run outside a React hook while keeping the same behavior used internally by the hooks.

```typescript
import {
  createScopedSPFxStorageKey,
  createSPFxPnPListService,
  getSPFxUserInfo,
} from '@apvee/spfx-react-toolkit';

const user = getSPFxUserInfo(pageContext);
const filtersKey = createScopedSPFxStorageKey(instanceId, 'filters');
const tasks = createSPFxPnPListService(sp, 'Tasks', 50);
```

### API Permission Precheck

Use `useSPFxApiPermissionPrecheck` to check whether the current SPFx runtime can obtain delegated tokens for Microsoft Graph and custom APIs before running a feature that depends on those scopes.

```typescript
import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';

function PermissionStatus() {
  const precheck = useSPFxApiPermissionPrecheck({
    graph: ['Sites.Read.All'],
    customApis: [
      {
        name: 'Orders API',
        resource: 'api://contoso-orders-api',
        packageResource: 'Orders API',
        scopes: ['Orders.Read']
      }
    ]
  });

  return <span>{precheck.configurationState}</span>;
}
```

`available` means the current SPFx runtime obtained a token with the delegated scope. It does not read tenant grants and does not replace server-side authorization.

---

## Development Scripts

Use these scripts when working on this repository:

| Script | Purpose |
|--------|---------|
| `npm run build` | Bundles the SPFx package with `gulp bundle`, including TypeScript, Sass, lint, and webpack steps. |
| `npm run clean` | Removes generated SPFx build output. Run before a clean build or before publishing. |
| `npm test` | Runs the SPFx test pipeline with `gulp test`. |
| `npm run verify:examples` | Verifies that the sample webpart registry covers all exported hooks and providers. |
| `npm run verify:runtime-store` | Runs a focused runtime-store behavior check without requiring SharePoint. |
| `npm run verify:public-docs` | Verifies that public helper/service documentation stays aligned with exported APIs. |
| `npm run prepublishOnly` | Runs automatically before `npm publish`; currently performs a clean build. |

Recommended local check before opening a PR or publishing a package:

```bash
npm run clean
npm run build
npm test
npm run verify:examples
npm run verify:runtime-store
npm run verify:public-docs
```

---

## 📚 Documentation

For complete documentation including:
- Installation & configuration
- All 4 provider components
- Complete hooks API reference (40 hooks)
- Public helpers and services API references
- Code examples and best practices

**➡️ [View Full Documentation](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INTRODUCTION.md)**

---

## Requirements

| Requirement | Version |
|-------------|---------|
| Node.js | 22.x |
| SPFx | 1.18.0+ |
| React | 17.x |
| TypeScript | 5.3+ |

---

## License

MIT — See [LICENSE](./LICENSE) for details.

---

## Links

- [📖 Full Documentation](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INTRODUCTION.md)
- [📚 API Reference](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/INDEX.md)
- [🧰 Helpers API](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/helpers/INDEX.md)
- [🧩 Services API](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/api/services/INDEX.md)
- [📦 NPM Package](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)
- [🐛 Issues](https://github.com/apvee/spfx-react-toolkit/issues)

---

Made with ❤️ by [Apvee Solutions](https://github.com/apvee)
