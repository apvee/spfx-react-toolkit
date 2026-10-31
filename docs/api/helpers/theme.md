# Theme Helpers

[Helpers overview](../INDEX.md#helpers-and-styles) · [Non-React composition](../../guides/helpers-and-services.md)

Import:

```ts
import {
  createFluent9ThemeFromSPFxTheme,
  getTeamsFluentTheme,
} from '@apvee/spfx-react-toolkit';
```

## `createFluent9ThemeFromSPFxTheme`

```ts
function createFluent9ThemeFromSPFxTheme(spfxTheme: IReadonlyTheme | undefined): Theme;
```

Converts an SPFx Fluent UI 8 theme to a Fluent UI 9 theme. When `spfxTheme` is undefined, returns `webLightTheme`.

The existing Fluent migration shim supplies the conversion. Two accessible neutral stroke state tokens use the host palette so interactions remain observable:

| Fluent UI 9 token | SPFx palette source |
| --- | --- |
| `colorNeutralStrokeAccessible` | Existing shim mapping (`neutralSecondary`) |
| `colorNeutralStrokeAccessibleHover` | `neutralPrimary` |
| `colorNeutralStrokeAccessiblePressed` | `neutralDark` |

Every other converted token keeps the shim's mapping. Missing, empty or whitespace-only hover/pressed palette values retain the corresponding converted fallback. The helper does not mutate the input theme. Custom/inverted palettes use their supplied colors; palettes with equal colors can still have equal state tokens and require their own contrast/feedback checks. Teams themes returned by `getTeamsFluentTheme` remain the original shared theme objects. The existing Fluent theme/migration peers supply the conversion; no extra dependency is required.

```tsx
const theme = createFluent9ThemeFromSPFxTheme(spfxTheme);
return <FluentProvider theme={theme}>{children}</FluentProvider>;
```

For example, the default Fluent UI 8 host palette provides `#605e5c` for the base accessible stroke, `#323130` for hover and `#201f1e` for pressed. [`scrollbar.fluent`](styles.md#overflow-and-the-single-scrollbar-recipe) consumes these scoped tokens for its thumb feedback. `useSPFxFluent9ThemeInfo` uses this conversion in SharePoint and keeps its existing Teams/fallback selection order.

## `getTeamsFluentTheme`

```ts
function getTeamsFluentTheme(teamsThemeName: string | undefined): Theme;
```

| Teams theme name | Fluent UI 9 theme |
|------------------|-------------------|
| `dark` | `teamsDarkTheme` |
| `contrast`, `highcontrast` | `teamsHighContrastTheme` |
| `default`, undefined, other values | `teamsLightTheme` |

`getTeamsFluentTheme` lowercases the supplied name before matching, so `DARK` and `HighContrast` select their respective shared themes. It does not trim surrounding whitespace: ` dark ` falls back to Teams light. The helper returns the original shared Fluent theme objects; treat them and the undefined-SPFx-theme light fallback as immutable. Neither function mounts a provider, inserts theme variables or loads host context.

```ts
const light = createFluent9ThemeFromSPFxTheme(undefined);
const teamsDark = getTeamsFluentTheme('DARK');
const fallback = getTeamsFluentTheme(undefined);
```

Supply the returned `Theme` to a Fluent provider to establish variables. The [theming hooks](../hooks/theming.md) additionally obtain host data and subscribe to changes; these functions only convert supplied values.
