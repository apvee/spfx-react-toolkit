# Compose typed styles

[Documentation home](../INDEX.md) · [Complete style API](../api/helpers/styles.md)

Use `useSx` inside a React component to compose descriptors into a CSS class:

```tsx
import * as React from 'react';
import { useSx, flex, gap, padding, presets } from '@apvee/spfx-react-toolkit/styles';

export function Summary(): React.ReactElement {
  const sx = useSx();
  return <div className={sx(flex.row, gap.medium, padding.medium, presets.canvas)}>Summary</div>;
}
```

Use the host's shared React/Griffel runtime and renderer. Token-based selections need theme variables in scope, normally through the existing FluentProvider; the hook needs no SPFx provider and does not create a new renderer. See [installation](../getting-started/installation.md#shared-runtimes).

The [reference](../api/helpers/styles.md) owns the qualified descriptor catalog, composition order, state/query precedence, theme limitations and cache lifetime. The [generic hook reference](../api/hooks/react.md#usesx) documents hook options. Use those contracts for responsive, state and scrollbar recipes instead of copying mappings into this guide.

The lazy Styles sample provides an interactive catalog; [local browser checks](../development/browser-checks.md) and [authenticated host checks](../development/sharepoint-validation/checklists/styling.md) establish different verification boundaries.
