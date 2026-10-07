# Styles and useSx

Style descriptors are immutable, typed data used to describe dimensions, Fluent theme references, typography, and CSS scopes. They are independent of SPFx context. Creating or importing a descriptor does not access the DOM or insert CSS.

The complete style catalog is available from the package root, the clean `@apvee/spfx-react-toolkit/styles` facade and the historical `@apvee/spfx-react-toolkit/lib/helpers/styles` deep entry point. The facade also re-exports canonical `useSx` and all public descriptor types; the narrow `/styles/useSx` entry exposes only that hook. Root, facade and historical leaves compose with the same active renderer/direction context. See [package imports](../../PACKAGE-IMPORTS.md) for resolver/peer contracts and bundle interpretation.

```tsx
import { useSx, width, typography, SxInput } from '@apvee/spfx-react-toolkit/styles';
import * as paddingInlineStart from '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start';
// In a component: const sx = useSx();
const inputs: readonly SxInput[] = [width.px(240), typography.body1, paddingInlineStart.px(16)];
```

Existing root imports remain valid:

```ts
import {
  width, minWidth, foreground, typography,
  responsive, viewport, hover, active, focusVisible,
  SxBaseDescriptor, SxInput, SxFunction, SxOptions,
} from '@apvee/spfx-react-toolkit';

const surface: readonly SxInput[] = [
  width.full,
  minWidth.zero,
  typography.body1,
  foreground.subtle,
  responsive.medium(hover(foreground.subtle)),
];
const fixedWidth: SxBaseDescriptor = width.px(240);
```

Descriptors produce a class name through `useSx`, which reads the existing Griffel renderer and Fluent direction context. The descriptors alone do not insert CSS. The concrete deep imports include the package's `lib` directory:

```ts
import { useSx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx';
import { presets, typography, padding } from '@apvee/spfx-react-toolkit/lib/helpers/styles';

// Inside a component:
const sx = useSx();
const className = sx(presets.canvas, typography.body1, padding.medium);
```

## Basic API examples

| Namespace or function | Selection | Meaning |
| --- | --- | --- |
| `width` | `auto`, `full`, `px(value)` | `auto`, `100%`, or an explicit pixel length |
| `minWidth` | `zero`, `px(value)` | `0px` or an explicit pixel length |
| `foreground` | `subtle` | `var(--colorNeutralForeground2)` |
| `typography` | `body1` | Official Fluent font family, size, weight, and line height |
| `hover`, `active`, `focusVisible` | `(...baseDescriptors)` | One CSS pseudo-state |
| `responsive` | `small`, `medium`, `large` | Container inline-size thresholds of 480, 640, and 1024px |
| `viewport` | `small`, `medium`, `large` | Viewport width thresholds of 480, 640, and 1024px |

All catalog pixel factories accept finite nonnegative numbers, including zero, and preserve explicit `px` units. Negative values, `NaN`, and infinity throw `RangeError`; values are never silently clamped.

### Fluent token references

The library source uses `tokens` from `@fluentui/react-theme` for token-backed declarations. These references contain CSS variable expressions, not resolved theme colors:

```ts
import { tokens } from '@fluentui/react-theme';

// Library source value used by foreground.subtle:
const subtleColor = tokens.colorNeutralForeground2;
// Value emitted for the color declaration: var(--colorNeutralForeground2)
```

Consumers continue to use `sx(foreground.subtle)`. The browser resolves the emitted CSS variable against the element's scoped Fluent theme, so theme changes preserve the same descriptor API and theme scoping. The CSS variable mappings below describe emitted values.

`body1` uses the installed `@fluentui/react-theme` recipe. Its values remain CSS variable references: `fontFamilyBase`, `fontSizeBase300`, `fontWeightRegular`, and `lineHeightBase300`. It does not set color. Theme variables must be supplied in the element's scope, usually by a Fluent provider; descriptor construction itself requires no provider.

## Types and scopes

```ts
type SxDescriptor = SxBaseDescriptor | SxStateDescriptor | SxResponsiveDescriptor;
type SxInput = SxDescriptor | string | false | null | undefined;
type SxFunction = (...inputs: readonly SxInput[]) => string;
interface SxOptions { readonly dir?: 'ltr' | 'rtl'; }
```

All three descriptor types are readonly and opaque. Raw CSS objects do not satisfy them. States accept only base descriptors; responsive and viewport functions accept base or state descriptors. The canonical combined form is `responsive.medium(hover(foreground.subtle))`. Nested states, nested queries, and class strings inside scopes are rejected by TypeScript. Existing class strings and conditional falsy values belong at the composer level.

Container queries target the nearest ancestor named `apvee-sx`; their thresholds are inclusive. The element cannot query its own inline-size. Viewport queries use the window width. Scope constructors preserve data only; the browser evaluates the CSS materialized by `sx`.

`SxOptions.dir` overrides the effective `ltr` or `rtl` direction. When omitted, `useSx` uses Fluent direction context, with LTR as the context-free default. Keep it consistent with the host's direction and any external Griffel classes when using a standalone Griffel context.


## Composition and precedence

Call `useSx` once at the top level of a React component; its returned `sx` is an ordinary function and can be called repeatedly while rendering. It returns a `className` string and never requires an inline `style` attribute. Its identity is stable while the Griffel renderer and effective direction stay unchanged. `sx()` and inputs consisting only of `false`, `null`, `undefined`, or empty strings produce `''`.

Within one normalized property and the same scope, the last declaration wins. Recipes expand before resolving overrides, so `sx(padding.medium, paddingInlineStart.small)` replaces only inline-start. Different scopes have explicit priority independent of descriptor order or mount order:

1. `focusVisible` overrides `active`, which overrides `hover`, which overrides the base state for overlapping properties.
2. Within each state, matching `large` overrides `medium`, then `small`, then the base declaration.
3. At the same breakpoint within the same state, a matching container query overrides a matching viewport query.

State priority applies before breakpoint priority: an applicable hover base declaration wins over a non-state large query. Independent properties remain independent; omitting color allows ordinary CSS inheritance. Focus-visible styling does not itself create an outline or accessible focus indicator; retain a native outline or supply an appropriate control treatment.

Class strings keep their original position between contiguous descriptor segments. External Griffel atomic classes participate in `mergeClasses`: `sx(width.px(240), externalWidth400, gap.small)` retains the external 400px width; adding `width.px(360)` last changes width to 360px. Existing `sx` results can also be composed as class strings, preserving independent scopes. Keep one shared Griffel installation in the host. Ordinary CSS class strings are preserved and follow native specificity, stylesheet order and importance; argument order does not guarantee their property precedence.

Application states use normal conditions, rather than new CSS scopes. Remove enabled recipes when disabled; a disabled base preset cannot cancel a more specific enabled hover recipe:

```tsx
const sx = useSx();
return <button
  disabled={disabled}
  className={sx(
    disabled ? presets.disabled : selected ? presets.brandTint : presets.canvas,
    !disabled && hover(background.alternative),
    !disabled && active(background.brandTint),
    padding.medium,
  )}
>Action</button>;
```

## Container, viewport and direction

`container.inlineSize` sets exactly `containerType: inline-size` and `containerName: apvee-sx`. Apply it to an ancestor of the queried element. It adds no width, height or display; inline-size containment can affect intrinsic sizing, so provide suitable host geometry. The nearest named ancestor supplies the query size. Inline-size follows writing mode: horizontal writing measures width, vertical writing measures height. An element never queries itself; without a matching ancestor the base styles still apply.

```tsx
const sx = useSx();
return <section className={sx(container.inlineSize, width.full)}>
  <div className={sx(flex.column, responsive.medium(flex.row), gap.medium)}>
    <span>First</span><span>Second</span>
  </div>
</section>;
```

`responsive.small`, `responsive.medium` and `responsive.large` query inclusive ancestor inline-size thresholds of 480px, 640px and 1024px. `viewport.small`, `viewport.medium` and `viewport.large` use those inclusive window-width thresholds independently. Changing a container does not change the viewport query.

The hook reads the existing Griffel renderer and Fluent direction context. `SxOptions.dir` explicitly overrides that direction; without an override or different context, the default is LTR. In a standalone Griffel environment, keep `useSx({ dir: 'rtl' })`, the host's DOM/CSS direction, Fluent direction context and external Griffel classes coherent. The option does not set the DOM `dir` attribute. Logical spacing/alignment follows CSS direction and writing mode; do not assume physical left/right mappings in vertical layouts.

## Dependencies and cache lifetime

The package's mandatory shared style peers are `@griffel/core@^1.19.2`, `@griffel/react@^1.5.30` and `@fluentui/react-shared-contexts@^9.25.2`. Existing Fluent peers remain `@fluentui/react-migration-v8-v9@^9.9.12`, `@fluentui/react-theme@^9.2.0` and `@fluentui/react-utilities@^9.25.1`. The consuming host supplies compatible shared installations; the hook needs no SPFx provider and creates no renderer.

Token styles need Fluent theme CSS variables in the element's scope. A host can supply them through its existing FluentProvider. The sample uses `@fluentui/react-provider@9.22.8`; that provider package is optional for consumers and is not an additional toolkit peer. Switching scoped theme variables updates computed colors and sizes without RGB-dependent classes.

Factories and generated assignments are cached by renderer and canonical property/scope/value. Repeated values reuse classes and rules. A finite repeated set of values avoids continued growth from those values; continuously distinct pixel values still grow the assignment maps and Griffel CSS caches for the renderer's lifetime. There is no hard cache limit or CSS removal guarantee on component unmount. The renderer WeakMap does not bound memory while a renderer remains live. Prefer a finite design scale for frequently changing dimensions; do not generate a new arbitrary value every frame.

Factories are isolated per renderer to avoid capturing another renderer's `classNameHashSalt`. The installed Griffel core 1.19.2 emits its upstream development console diagnostic for a nonempty salt during the first factory computation, including an isolated factory. The toolkit preserves that diagnostic. Reuse the host renderer/salt configuration and investigate console output; this behavior is not suppressed or presented as a warning-free custom-salt path.

## Complete v1 catalog mappings

Each namespace consists of individual ESM exports returning opaque base descriptors, except the numeric factories and scope functions. The catalog references `@fluentui/react-theme` 9.2.0 and its actual `@fluentui/tokens` 1.0.0-alpha.22 floor. Token selections retain `var(--tokenName)` references, so changing a scoped Fluent theme updates computed styles without generating RGB-dependent classes. A scoped Fluent theme, normally from `FluentProvider`, is required for token-based values; the generic hook requires no SPFx context.

### Layout and dimensions

| Namespace | Selections and CSS mapping |
| --- | --- |
| `flex` | `row` / `column`: `display: flex` and `flexDirection: row` / `column`; `wrap` / `noWrap`: `flexWrap: wrap` / `nowrap` |
| `flexItem` | `grow` / `noGrow`: `flexGrow: 1` / `0`; `shrink` / `noShrink`: `flexShrink: 1` / `0` |
| `grid` | `columns(count)`: `display: grid`, `gridTemplateColumns: repeat(count, minmax(0, 1fr))`; positive integers only |
| `alignItems` | `start`, `center`, `end`, `stretch`, `baseline` map to their CSS keywords |
| `justifyContent` | `start`, `center`, `end`, `spaceBetween` map to `start`, `center`, `end`, `space-between` |
| `alignSelf` | `auto`, `start`, `center`, `end`, `stretch` map to their CSS keywords |
| `width`, `height` | `auto`: `auto`; `full`: `100%`; `px(value)`: explicit finite nonnegative pixels |
| `minWidth`, `minHeight` | `zero`: `0px`; `px(value)`: explicit finite nonnegative pixels |
| `maxWidth`, `maxHeight` | `none`: `none`; `px(value)`: explicit finite nonnegative pixels |

Alignment and gap do not change display. Logical `start` and `end` follow CSS writing mode and direction. Percentage height needs a definite containing-block height. Pixel factories preserve fractions and zero; negative, NaN, or infinite values throw `RangeError`. `grid.columns` rejects zero, fractions, negative, NaN, or infinite counts with `RangeError`.

### Spacing

All 15 spacing namespaces export `none`, `extraSmall`, `small`, `medium`, `large`, `extraLarge`, and `px(value)`. The names map to token suffixes `None`, `XS`, `S`, `M`, `L`, and `XL`. Block sides use `spacingVertical*`, inline sides use `spacingHorizontal*`; values remain CSS variables. Pixel factories apply the same explicit pixel length to the selected longhands.

| Namespace | Normalized properties |
| --- | --- |
| `gap` | `rowGap` using vertical tokens; `columnGap` using horizontal tokens |
| `padding`, `margin` | Both block and both inline longhands |
| `paddingInline`, `marginInline` | Inline start and inline end |
| `paddingBlock`, `marginBlock` | Block start and block end |
| `paddingInlineStart`, `marginInlineStart` | Inline start only |
| `paddingInlineEnd`, `marginInlineEnd` | Inline end only |
| `paddingBlockStart`, `marginBlockStart` | Block start only |
| `paddingBlockEnd`, `marginBlockEnd` | Block end only |

For example, `sx(padding.medium, paddingInlineStart.small)` changes only inline-start to `spacingHorizontalS`; the other three medium sides remain. Margin pixel factories follow the same nonnegative contract as padding.

### Foreground, background, and presets

The following token names include their full `color` prefix. Foreground selections set only `color`; background selections set only `backgroundColor`. Neither adds geometry, semantics, or interaction states. `link` adds color without link behavior; `disabled` adds styling without disabling an element.

| Foreground selection | Fluent token |
| --- | --- |
| `primary` | `colorNeutralForeground1` |
| `subtle` | `colorNeutralForeground2` |
| `muted` | `colorNeutralForeground3` |
| `disabled` | `colorNeutralForegroundDisabled` |
| `brand` | `colorBrandForeground1` |
| `onBrand` | `colorNeutralForegroundOnBrand` |
| `inverted` | `colorNeutralForegroundInverted` |
| `link` | `colorBrandForegroundLink` |

| Background selection | Fluent token |
| --- | --- |
| `canvas` | `colorNeutralBackground1` |
| `alternative` | `colorNeutralBackground2` |
| `subtle` | `colorSubtleBackground` |
| `transparent` | `colorTransparentBackground` |
| `brand` | `colorBrandBackground` |
| `brandTint` | `colorBrandBackground2` |
| `inverted` | `colorNeutralBackgroundInverted` |
| `disabled` | `colorNeutralBackgroundDisabled` |

`subtle` and `alternative` remain distinct roles. The floor theme's subtle background is transparent at rest. `subtle` and `transparent` remain distinct tokens even when their resolved values coincide.

| Preset | Background token | Foreground token |
| --- | --- | --- |
| `canvas` | `colorNeutralBackground1` | `colorNeutralForeground1` |
| `alternative` | `colorNeutralBackground2` | `colorNeutralForeground1` |
| `subtle` | `colorSubtleBackground` | `colorNeutralForeground2` |
| `transparent` | `colorTransparentBackground` | Omitted; ordinary color inheritance |
| `brand` | `colorBrandBackground` | `colorNeutralForegroundOnBrand` |
| `brandTint` | `colorBrandBackground2` | `colorBrandForeground2` |
| `inverted` | `colorNeutralBackgroundInverted` | `colorNeutralForegroundInverted` |
| `success` | `colorStatusSuccessBackground1` | `colorStatusSuccessForeground1` |
| `warning` | `colorStatusWarningBackground1` | `colorStatusWarningForeground1` |
| `danger` | `colorStatusDangerBackground1` | `colorStatusDangerForeground1` |
| `disabled` | `colorNeutralBackgroundDisabled` | `colorNeutralForegroundDisabled` |

Presets expand property by property. `sx(presets.canvas, foreground.subtle)` preserves the canvas background and replaces only color. Reversing those arguments restores the canvas foreground. Presets add no display, padding, dimensions, border, radius, shadow, or automatic Hover/Pressed variants.

The installed `teamsHighContrastTheme` maps both inverted tokens to black (1:1 contrast). **`presets.inverted` is supported for the verified light/dark themes; use `canvas` or `alternative` for content in that floor high contrast theme.** `presets.disabled` is restricted to actual disabled presentation; it does not set the native disabled attribute or `aria-disabled`. Its light/dark foreground pairs have low contrast. Transparent/subtle presets and scrollbar tracks depend on the underlying surface. Custom themes and backgrounds require their own contrast evaluation; the token names do not guarantee universal contrast.

The local Chrome 152 checkpoint measured all 11 presets on canvas and alternative surfaces in the three floor themes. Supported active-content pairs ranged from 4.85:1 to 21:1. Disabled pairs measured 1.65:1 in light and 2.76:1 in dark; inverted measured 1:1 in the Fluent high contrast theme. These measurements cover those exact palettes and surfaces, including the actual underlying color for transparent backgrounds.

### Typography and text

All 17 typography selections use the official recipe's four properties: `fontFamily`, `fontSize`, `fontWeight`, and `lineHeight`. Each uses `fontFamilyBase`; the table gives the corresponding size, line-height, and weight tokens. Typography leaves foreground inherited and does not assign HTML heading semantics.

| Selection | Font size token | Line height token | Weight token |
| --- | --- | --- | --- |
| `body1` | `fontSizeBase300` | `lineHeightBase300` | `fontWeightRegular` |
| `body1Strong` | `fontSizeBase300` | `lineHeightBase300` | `fontWeightSemibold` |
| `body1Stronger` | `fontSizeBase300` | `lineHeightBase300` | `fontWeightBold` |
| `body2` | `fontSizeBase400` | `lineHeightBase400` | `fontWeightRegular` |
| `caption1` | `fontSizeBase200` | `lineHeightBase200` | `fontWeightRegular` |
| `caption1Strong` | `fontSizeBase200` | `lineHeightBase200` | `fontWeightSemibold` |
| `caption1Stronger` | `fontSizeBase200` | `lineHeightBase200` | `fontWeightBold` |
| `caption2` | `fontSizeBase100` | `lineHeightBase100` | `fontWeightRegular` |
| `caption2Strong` | `fontSizeBase100` | `lineHeightBase100` | `fontWeightSemibold` |
| `subtitle1` | `fontSizeBase500` | `lineHeightBase500` | `fontWeightSemibold` |
| `subtitle2` | `fontSizeBase400` | `lineHeightBase400` | `fontWeightSemibold` |
| `subtitle2Stronger` | `fontSizeBase400` | `lineHeightBase400` | `fontWeightBold` |
| `title1` | `fontSizeHero800` | `lineHeightHero800` | `fontWeightSemibold` |
| `title2` | `fontSizeHero700` | `lineHeightHero700` | `fontWeightSemibold` |
| `title3` | `fontSizeBase600` | `lineHeightBase600` | `fontWeightSemibold` |
| `largeTitle` | `fontSizeHero900` | `lineHeightHero900` | `fontWeightSemibold` |
| `display` | `fontSizeHero1000` | `lineHeightHero1000` | `fontWeightSemibold` |

`textAlign.start/center/end` set their CSS alignment keywords. `text.wrap` sets `whiteSpace: normal`; `text.noWrap` sets `whiteSpace: nowrap`. `text.truncate` adds `whiteSpace: nowrap`, `overflowX/Y: hidden`, and `textOverflow: ellipsis`. Truncation needs a constrained available width; it does not add display or minimum width.

### Borders, radius, and shadow

| Namespace | Mapping |
| --- | --- |
| `borderWidth` | `none`: `0px`; `thin`: `strokeWidthThin`; `thick`: `strokeWidthThick`, on all four width longhands |
| `borderStyle` | `solid`: `solid`, on all four style longhands |
| `borderColor` | `primary`: `colorNeutralStroke1`; `subtle`: `colorNeutralStroke2`; `brand`: `colorBrandStroke1`; `disabled`: `colorNeutralStrokeDisabled`, on all four color longhands |
| `borderRadius` | `none/small/medium/large/extraLarge/circular`: `borderRadiusNone/Small/Medium/Large/XLarge/Circular`, on all four corners |
| `boxShadow` | `none`: `none`; `small`: `shadow4`; `medium`: `shadow8` |

Border geometry remains explicit: compose `borderWidth.thin`, `borderStyle.solid`, and `borderColor.primary` when a visible border is required. Token-based decoration requires theme variables in scope.

### Overflow and the single scrollbar recipe

`overflow.visible/hidden/auto` expand to both `overflowX` and `overflowY`. `overflow.horizontal.visible/hidden/auto` set only `overflowX`; `overflow.vertical.visible/hidden/auto` set only `overflowY`. The browser retains native axis coupling: a visible axis can compute to auto when the other axis is hidden or auto.

`scrollbar.fluent` is the only scrollbar selection (`SxBaseDescriptor`). In Edge/Chrome and other browsers supporting `@supports selector(::-webkit-scrollbar)`, both axes use **6px** scrollbars with a `colorNeutralStrokeAccessible` thumb and transparent track/corner. These pseudo-element rules apply only under `(forced-colors: none)`; standard `scrollbarWidth` and `scrollbarColor` become `auto` in that branch so they do not suppress vendor styling. Other browsers retain standard `scrollbarWidth: thin` and `scrollbarColor: var(--colorNeutralStrokeAccessible) transparent`. An actual `(forced-colors: active)` media rule sets `scrollbarColor: auto`, preserving native thin sizing and browser system colors. The recipe adds no overflow, container dimensions, gutter, overscroll, smooth scrolling, or `forced-color-adjust: none`.

The vendor thumb retains **3px rounded corners** on both axes. Its base color is `colorNeutralStrokeAccessible`; moving the pointer over the thumb itself uses `colorNeutralStrokeAccessibleHover`, and pressing or dragging it uses `colorNeutralStrokeAccessiblePressed`. Pressed feedback wins while the thumb is also hovered; release restores hover feedback and leaving restores the base color. These tokens follow the existing Fluent theme variables in scope. Hovering the content or track does not activate thumb feedback. Forced colors and browsers using the standard fallback retain native interactions.

```tsx
const sx = useSx();
return <div className={sx(width.px(320), height.px(160), overflow.auto, scrollbar.fluent, presets.canvas)}>
  <div className={sx(width.px(600), height.px(400))}>Content scrolls on both axes.</div>
</div>;
```

Use the existing shared Griffel/Fluent peers and scoped theme variables described above; no extra dependency is required. Browsers without vendor pseudo-element support need standard `scrollbar-width` and `scrollbar-color` for the thin fallback; otherwise native appearance remains. Platform overlay scrollbars may appear only during scrolling. `hover`, `active`, `focusVisible`, `responsive` and `viewport` scopes apply the vendor rules only while their scopes are active; removing the descriptor restores native appearance. The transparent track reveals the actual underlying surface, so thumb contrast must be checked there.

Foreign Griffel declarations retain the usual argument/merge order: a later native `scrollbarWidth: none` hides the scrollbar, and a later `scrollbarColor` overrides its colors. In modern Chromium, non-`auto` standard scrollbar values select native rendering and can bypass the recipe's vendor 6px sizing. Putting `scrollbar.fluent` last restores the recipe in the same scope.


The same local Chrome checkpoint measured the accessible neutral scrollbar thumb against the tested underlying surfaces:

| Floor theme | Canvas | Alternative |
| --- | --- | --- |
| `webLightTheme` | 6.19:1 | 5.93:1 |
| `webDarkTheme` | 6.48:1 | 7.34:1 |
| `teamsHighContrastTheme` | 21:1 | 21:1 |

Browser forced colors were emulated separately from the Fluent high contrast palette and computed `scrollbarColor: auto` in each theme. Local browser measurements do not establish authenticated SharePoint-host behavior.

## Qualified API reference

Base selections below have type `SxBaseDescriptor`; numeric factories return the same opaque type. State and query factories return their respective scope descriptor types, with signatures given after the table. Each qualified member is listed explicitly so IntelliSense, documentation and export checks refer to the same public surface.

| Namespace | Public selections and factories |
| --- | --- |
| `width` | `width.auto`, `width.full`, `width.px(value: number)` |
| `minWidth` | `minWidth.zero`, `minWidth.px(value: number)` |
| `foreground` | `foreground.primary`, `foreground.subtle`, `foreground.muted`, `foreground.disabled`, `foreground.brand`, `foreground.onBrand`, `foreground.inverted`, `foreground.link` |
| `typography` | `typography.body1`, `typography.body1Strong`, `typography.body1Stronger`, `typography.body2`, `typography.caption1`, `typography.caption1Strong`, `typography.caption1Stronger`, `typography.caption2`, `typography.caption2Strong`, `typography.subtitle1`, `typography.subtitle2`, `typography.subtitle2Stronger`, `typography.title1`, `typography.title2`, `typography.title3`, `typography.largeTitle`, `typography.display` |
| `responsive` | `responsive.small(...inputs)`, `responsive.medium(...inputs)`, `responsive.large(...inputs)` |
| `viewport` | `viewport.small(...inputs)`, `viewport.medium(...inputs)`, `viewport.large(...inputs)` |
| `container` | `container.inlineSize` |
| `height` | `height.auto`, `height.full`, `height.px(value: number)` |
| `minHeight` | `minHeight.zero`, `minHeight.px(value: number)` |
| `maxWidth` | `maxWidth.none`, `maxWidth.px(value: number)` |
| `maxHeight` | `maxHeight.none`, `maxHeight.px(value: number)` |
| `flex` | `flex.row`, `flex.column`, `flex.wrap`, `flex.noWrap` |
| `flexItem` | `flexItem.grow`, `flexItem.noGrow`, `flexItem.shrink`, `flexItem.noShrink` |
| `grid` | `grid.columns(count: number)` |
| `alignItems` | `alignItems.start`, `alignItems.center`, `alignItems.end`, `alignItems.stretch`, `alignItems.baseline` |
| `justifyContent` | `justifyContent.start`, `justifyContent.center`, `justifyContent.end`, `justifyContent.spaceBetween` |
| `alignSelf` | `alignSelf.auto`, `alignSelf.start`, `alignSelf.center`, `alignSelf.end`, `alignSelf.stretch` |
| `gap` | `gap.none`, `gap.extraSmall`, `gap.small`, `gap.medium`, `gap.large`, `gap.extraLarge`, `gap.px(value: number)` |
| `padding` | `padding.none`, `padding.extraSmall`, `padding.small`, `padding.medium`, `padding.large`, `padding.extraLarge`, `padding.px(value: number)` |
| `paddingInline` | `paddingInline.none`, `paddingInline.extraSmall`, `paddingInline.small`, `paddingInline.medium`, `paddingInline.large`, `paddingInline.extraLarge`, `paddingInline.px(value: number)` |
| `paddingBlock` | `paddingBlock.none`, `paddingBlock.extraSmall`, `paddingBlock.small`, `paddingBlock.medium`, `paddingBlock.large`, `paddingBlock.extraLarge`, `paddingBlock.px(value: number)` |
| `paddingInlineStart` | `paddingInlineStart.none`, `paddingInlineStart.extraSmall`, `paddingInlineStart.small`, `paddingInlineStart.medium`, `paddingInlineStart.large`, `paddingInlineStart.extraLarge`, `paddingInlineStart.px(value: number)` |
| `paddingInlineEnd` | `paddingInlineEnd.none`, `paddingInlineEnd.extraSmall`, `paddingInlineEnd.small`, `paddingInlineEnd.medium`, `paddingInlineEnd.large`, `paddingInlineEnd.extraLarge`, `paddingInlineEnd.px(value: number)` |
| `paddingBlockStart` | `paddingBlockStart.none`, `paddingBlockStart.extraSmall`, `paddingBlockStart.small`, `paddingBlockStart.medium`, `paddingBlockStart.large`, `paddingBlockStart.extraLarge`, `paddingBlockStart.px(value: number)` |
| `paddingBlockEnd` | `paddingBlockEnd.none`, `paddingBlockEnd.extraSmall`, `paddingBlockEnd.small`, `paddingBlockEnd.medium`, `paddingBlockEnd.large`, `paddingBlockEnd.extraLarge`, `paddingBlockEnd.px(value: number)` |
| `margin` | `margin.none`, `margin.extraSmall`, `margin.small`, `margin.medium`, `margin.large`, `margin.extraLarge`, `margin.px(value: number)` |
| `marginInline` | `marginInline.none`, `marginInline.extraSmall`, `marginInline.small`, `marginInline.medium`, `marginInline.large`, `marginInline.extraLarge`, `marginInline.px(value: number)` |
| `marginBlock` | `marginBlock.none`, `marginBlock.extraSmall`, `marginBlock.small`, `marginBlock.medium`, `marginBlock.large`, `marginBlock.extraLarge`, `marginBlock.px(value: number)` |
| `marginInlineStart` | `marginInlineStart.none`, `marginInlineStart.extraSmall`, `marginInlineStart.small`, `marginInlineStart.medium`, `marginInlineStart.large`, `marginInlineStart.extraLarge`, `marginInlineStart.px(value: number)` |
| `marginInlineEnd` | `marginInlineEnd.none`, `marginInlineEnd.extraSmall`, `marginInlineEnd.small`, `marginInlineEnd.medium`, `marginInlineEnd.large`, `marginInlineEnd.extraLarge`, `marginInlineEnd.px(value: number)` |
| `marginBlockStart` | `marginBlockStart.none`, `marginBlockStart.extraSmall`, `marginBlockStart.small`, `marginBlockStart.medium`, `marginBlockStart.large`, `marginBlockStart.extraLarge`, `marginBlockStart.px(value: number)` |
| `marginBlockEnd` | `marginBlockEnd.none`, `marginBlockEnd.extraSmall`, `marginBlockEnd.small`, `marginBlockEnd.medium`, `marginBlockEnd.large`, `marginBlockEnd.extraLarge`, `marginBlockEnd.px(value: number)` |
| `background` | `background.canvas`, `background.alternative`, `background.subtle`, `background.transparent`, `background.brand`, `background.brandTint`, `background.inverted`, `background.disabled` |
| `presets` | `presets.canvas`, `presets.alternative`, `presets.subtle`, `presets.transparent`, `presets.brand`, `presets.brandTint`, `presets.inverted`, `presets.success`, `presets.warning`, `presets.danger`, `presets.disabled` |
| `textAlign` | `textAlign.start`, `textAlign.center`, `textAlign.end` |
| `text` | `text.wrap`, `text.noWrap`, `text.truncate` |
| `borderWidth` | `borderWidth.none`, `borderWidth.thin`, `borderWidth.thick` |
| `borderStyle` | `borderStyle.solid` |
| `borderColor` | `borderColor.primary`, `borderColor.subtle`, `borderColor.brand`, `borderColor.disabled` |
| `borderRadius` | `borderRadius.none`, `borderRadius.small`, `borderRadius.medium`, `borderRadius.large`, `borderRadius.extraLarge`, `borderRadius.circular` |
| `boxShadow` | `boxShadow.none`, `boxShadow.small`, `boxShadow.medium` |
| `overflow.horizontal` | `overflow.horizontal.visible`, `overflow.horizontal.hidden`, `overflow.horizontal.auto` |
| `overflow.vertical` | `overflow.vertical.visible`, `overflow.vertical.hidden`, `overflow.vertical.auto` |
| `overflow` | `overflow.visible`, `overflow.hidden`, `overflow.auto` |
| `scrollbar` | `scrollbar.fluent` |

`hover(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor`, `active(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor`, and `focusVisible(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor` accept only base descriptors. Each `responsive.small`, `responsive.medium`, `responsive.large`, `viewport.small`, `viewport.medium`, and `viewport.large` function has signature `(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor`.

Every `.px(value: number): SxBaseDescriptor` factory accepts a finite nonnegative number (including fractions and zero), expressed in pixels. `grid.columns(count: number): SxBaseDescriptor` accepts finite positive integers. Invalid arguments throw `RangeError`. CSS layout limits still apply to large valid numbers.
