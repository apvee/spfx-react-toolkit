# Package imports and tree shaking

The package publishes ESNext JavaScript for a bundler and TypeScript declarations. Root imports remain supported; domain aliases offer a smaller, explicit API boundary. All aliases re-export the canonical implementations: mixed imports share hook identity, context objects, renderer access and descriptor semantics. They do not create extra provider instances or caches.

## Supported entry points

| Import suffix after `@apvee/spfx-react-toolkit` | Surface |
| --- | --- |
| (root) | All public providers, hooks, helpers and services |
| `/core` | Providers and public core types |
| `/hooks` | All public hooks |
| `/helpers` | Public helpers and descriptor namespaces |
| `/services` | Standalone service factories and types |
| `/styles` | All public descriptors, descriptor types and canonical `useSx` |
| `/styles/useSx` | Canonical `useSx` only |
| `/hooks/useStableCallback` | Canonical `useStableCallback` only |

Historical published `/lib/...` paths remain accessible, including extensionless modules, `.js` modules and directory aliases for historical `index` entries. Existing declarations, maps and physical paths remain published. Historical internal paths are compatibility paths, not newly recommended clean APIs. No wildcard alias grants access to arbitrary new internals.

```tsx
// Root API: existing imports continue to work.
import { useStableCallback, useSx, width } from '@apvee/spfx-react-toolkit';
// Domain alternatives (choose one hook binding per component).
import { useStableCallback as domainCallback } from '@apvee/spfx-react-toolkit/hooks';
import { useSx as domainSx, typography, SxInput } from '@apvee/spfx-react-toolkit/styles';
// Narrow clean hooks and preserved legacy leaves.
import { useSx as narrowSx } from '@apvee/spfx-react-toolkit/styles/useSx';
import { useStableCallback as narrowCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
import { useSx as legacySx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx';
import * as legacyWidth from '@apvee/spfx-react-toolkit/lib/helpers/styles/width';
```

Individual legacy descriptor modules export members such as `px` and `full`; their namespace name (`width`) is supplied by the barrel. Use a namespace import at the leaf, as above.

## What the Imports sample demonstrates

The web part's Imports panel exercises the same behavior through three import forms:

- **Root imports** use `@apvee/spfx-react-toolkit`.
- **Domain imports** use functional entry points such as `/hooks` and `/styles`.
- **Legacy leaf imports** use historical paths to individual modules, such as `/lib/hooks/useSx`. A leaf is the individual implementation module rather than a root or domain barrel.

The legacy card verifies compatibility for existing consumers after adding the new export map. Its captured callback must retain identity and read current state, and its styles must compose with descriptors imported through the root and domain entries beneath the same FluentProvider. All three forms resolve to the canonical implementations and contexts. The card is a compatibility scenario; new code can use the root, domain entries or clean narrow aliases such as `/styles/useSx`.

The panel makes runtime compatibility observable. The separate installed-package bundle gate establishes tree shaking; the full demo bundle is not a minimal-consumer measurement.

## Styles facade and generic callbacks

`/styles` re-exports `SxBaseDescriptor`, `SxStateDescriptor`, `SxResponsiveDescriptor`, `SxDescriptor`, `SxInput`, `SxFunction`, `SxOptions`, every public style namespace/scope, and `useSx`. It adds no descriptor implementation or registration. See the complete [style API](./api/helpers/styles.md) and [hook signatures](./api/hooks/react.md).

```tsx
import * as React from 'react';
import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
import { useSx, width, typography } from '@apvee/spfx-react-toolkit/styles';
import * as paddingInlineStart from '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start';

export function Counter() {
  const [count, setCount] = React.useState(0);
  const [message, setMessage] = React.useState('');
  // Argument, result and Promise types are inferred; no dependency array.
  const report = useStableCallback((label: string): string => `${label}: ${count}`);
  const sx = useSx();
  return <div className={sx(width.px(240), typography.body1, paddingInlineStart.px(16))}>
    <button onClick={() => setCount(value => value + 1)}>Increment</button>
    <button onClick={() => setMessage(report('Current count'))}>Report</button>
    <p>{message}</p>
  </div>;
}
```

Hooks must run unconditionally in stable component types. The callback reads committed state after its layout effect; do not invoke it during render. Descriptors remain immutable data and importing them inserts no CSS. `useSx` reads the active Griffel renderer and Fluent direction context; token recipes require theme CSS variables in scope. A FluentProvider is one way to supply them, not an additional mandatory peer solely for this hook. Shared peer instances are necessary for shared contexts. Descriptor order, state/query priority, native CSS limitations, theme/contrast limits and renderer-owned CSS/cache lifetime are unchanged. Arbitrary values can grow renderer-owned CSS; unmount does not guarantee CSS removal or a bounded cache.

## Resolver and peer requirements

Use the package through a bundler that understands `exports`, ESM and `sideEffects`. The verified consumer uses Webpack 5.95.0 and TypeScript 5.3.3 with `moduleResolution: node`; `typesVersions` supplies clean-alias declaration resolution for that TypeScript mode. Export conditions also expose `types` before the JavaScript `default` target. The sample compiles the library first and consumes its `lib` output, without source path aliases.

The tested baseline is Node 22.20.0, React/ReactDOM 17.0.1, SPFx 1.21.1 and PnPjs 4.17.0. Keep the declared SPFx, React, PnP, Griffel and Fluent peer contracts and share their instances with the host. Peer ranges are compatibility declarations, not evidence of a tested peer-version matrix. These additions do not provide native Node ESM execution, CommonJS output or a CommonJS runtime guarantee. Tests that execute a bundled Node probe verify that bundle and its adapter, not native package loading.

## PnP registration migration

Unrelated root imports no longer incidentally install all PnP features. Toolkit context/list/search/batch services register the features they own when used. Direct PnP callers, including `invoke` and `batch` callbacks, must import every feature their own operation uses:

```ts
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { createSPFxPnPService } from '@apvee/spfx-react-toolkit/services';
import type { SPFI } from '@pnp/sp';

export async function readTitles(sp: SPFI): Promise<unknown> {
  const service = createSPFxPnPService(sp);
  return service.invoke(client => client.web.lists.getByTitle('Tasks').items.select('Title')());
}
```

Files, folders and other features need their corresponding explicit PnP imports too. Registration does not configure authentication or grant permissions. See [PnP hooks](./api/hooks/pnpjs.md#pnp-feature-registrations) and the [service registration table](./api/services/INDEX.md#pnp-feature-registrations). Real SDK registrations are exercised locally with controlled transport; authenticated server permissions and caller-specific feature combinations require their own validation.

## What the bundle gate measures

The canonical production contract builds 36 variants and compares nine root/domain/legacy-leaf families. Blocking gzip gaps are at most 1,023 bytes for root/domain versus leaf, at most 512 bytes for the pure helper root versus leaf, and zero forbidden module groups or descriptor-retention violations. The verified informational snapshot currently has zero gaps across all nine families. Absolute raw/gzip/Brotli bytes are informational; compare emitted initial, async and union asset totals with matching compression/toolchain provenance. The checked-in snapshot is `tests/fixtures/tree-shaking/baseline.json`; exclusions and budgets live in `contract.json` next to it.

Style API reachability gaps and engine overhead answer different questions. The snapshot's minimal style case adds 1,732 gzip bytes versus direct Griffel and 1,396 bytes versus an already-Fluent control. This frozen snapshot predates the fixed-size scrollbar correction; current measurements are recorded separately in the maintenance verification rather than silently rewriting the baseline. These engine deltas are informational controls, not the alias gap budget or a universal app-size promise. A dynamic namespace lookup or a lazily loaded full catalog intentionally retains broader usage. The full lazy styles control emits an approximately 255 KiB async asset and a Webpack 244 KiB size advisory; it is not a static minimal-import regression. The full SPFx sample exercises a broad API surface and cannot establish the minimal-consumer tree-shaking result.

`npm run verify:package` installs the fresh tarball into an isolated real SPFx consumer, validates resolution, executes the canonical 36 bundle/18 runtime/9 mutation gates and builds the ship sample/solution. Transport-controlled runtime probes establish local registration, provider and renderer identity; they do not establish tenant authentication. Existing npm transitive deprecation/lifecycle warnings and SPFx metadata-age advisories remain visible. Check actual warning and failure output, including subprocess markers, instead of relying only on an outer success status.

See [development commands and baseline review](./DEVELOPMENT.md#production-import-gates-and-baseline-review) and the authenticated [Imports checklist](./SHAREPOINT-VALIDATION.md#imports).
