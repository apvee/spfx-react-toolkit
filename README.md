# SPFx React Toolkit

> A comprehensive React runtime and hooks library for SharePoint Framework (SPFx) with 35+ type-safe hooks. Simplifies SPFx development with instance-scoped state isolation and ergonomic hooks API across WebParts, Extensions, and Command Sets.

![SPFx React Toolkit](./assets/banner.png)

---

## Overview

**SPFx React Toolkit** is a production-ready library that simplifies SharePoint Framework development by providing a unified React context provider and a comprehensive collection of strongly-typed hooks.

Built on [Jotai](https://jotai.org/) atomic state management, it delivers per-instance state isolation, automatic synchronization, and an ergonomic React Hooks API that works across all SPFx component types.

### Key Benefits

| Benefit | Description |
|---------|-------------|
| 💪 **Type-Safe** | Full TypeScript support with zero `any` usage |
| ⚡ **Optimized** | Jotai atomic state with per-instance scoping (~3KB) |
| 🔄 **Auto-Sync** | Bidirectional synchronization between React and SPFx |
| 🎨 **Universal** | Works with WebParts, Application Customizers, Field Customizers, and Command Sets |
| 📦 **Modular** | Tree-shakeable, minimal bundle impact |

### Features

- ✅ **35+ React Hooks** — Comprehensive API surface for all SPFx capabilities
- ✅ **Instance Isolation** — State scoped per SPFx instance (multi-instance support)
- ✅ **PnPjs Integration** — Optional hooks for PnPjs v4 with type-safe filters
- ✅ **Cross-Platform** — Teams, SharePoint, and Local Workbench support

---

## Quick Start

### Installation

```bash
npm install @apvee/spfx-react-toolkit
```

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

---

## 📚 Documentation

For complete documentation including:
- Installation & configuration
- All 4 provider components
- Complete hooks API reference (35+ hooks)
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
- [📦 NPM Package](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)
- [🐛 Issues](https://github.com/apvee/spfx-react-toolkit/issues)

---

Made with ❤️ by [Apvee Solutions](https://github.com/apvee)
