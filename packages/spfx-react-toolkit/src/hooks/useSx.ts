/**
 * React styling utilities sharing the host Fluent direction and Griffel renderer.
 *
 * @packageDocumentation
 * @docPage hooks/react
 */
import { useMemo } from 'react';
import { useSxStyleEnvironment } from '../helpers/styles/renderer.internal';
import { resolveSxInputs } from '../helpers/styles/resolve.internal';
import type { SxFunction, SxOptions } from '../helpers/styles/types';

/**
 * Composes immutable style descriptors and existing classes in argument order using the current Griffel renderer and Fluent direction. No SPFx provider is required. Token descriptors need Fluent theme variables in the element's CSS scope. Empty/falsy composition returns an empty string. Last declaration wins per property/scope; external Griffel classes retain their position. Factories are cached per renderer and CSS rules persist for its lifetime.
 * @param options - Effective direction override; when omitted, uses Fluent context (LTR without context). Keep it consistent with host classes.
 * @returns Class composer, stable while renderer/direction stay unchanged.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width, foreground } from '@apvee/spfx-react-toolkit';
 *
 * function Surface() {
 *   const sx = useSx();
 *   return <div className={sx(width.px(240), foreground.subtle)}>Content</div>;
 * }
 * ```
 * @useWhen Compose toolkit descriptors and existing class names with the active host renderer.
 * @avoidWhen Continuously changing arbitrary pixel values would grow renderer caches; prefer a finite design scale.
 * @remarks
 * ### Scope and precedence
 * Call the hook at the top level of a component; call its ordinary returned composer repeatedly during render. Recipes expand into independent properties. Last declaration wins within one property/scope. State priority is focusVisible, active, hover, base; within a state, matching large, medium, small, base queries win in that order. At the same threshold, container overrides viewport independently of argument/mount order. A state declaration outranks a non-state query.
 * ### Existing classes
 * Class strings retain their original position between descriptor segments and participate in Griffel mergeClasses. Ordinary CSS classes retain native cascade/specificity rather than guaranteed argument-order precedence. Use one shared Griffel installation. Logical properties follow CSS writing mode/direction; the dir option does not set a DOM dir attribute.
 * ### Renderer and cache lifetime
 * The hook reads the existing renderer and Fluent direction and creates no renderer or SPFx provider. Token descriptors need scoped Fluent CSS variables, normally from the host FluentProvider. Factories/assignments are cached per renderer and canonical property/scope/value; they are not removed on component unmount and continuously distinct pixel values grow the cache for a live renderer. Repeated finite values reuse rules. Renderer-specific factories preserve classNameHashSalt isolation; upstream Griffel development salt diagnostics are retained.
 *
 *
 * @example Combine application state with CSS interaction scopes
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, hover, active, background, padding } from '@apvee/spfx-react-toolkit/styles';
 * function Action({ disabled, selected }: { disabled: boolean; selected: boolean }) {
 *   const sx = useSx();
 *   return <button disabled={disabled} className={sx(
 *     disabled ? presets.disabled : selected ? presets.brandTint : presets.canvas,
 *     !disabled && hover(background.alternative), !disabled && active(background.brandTint), padding.medium,
 *   )}>Action</button>;
 * }
 * ```
 *
 *
 * @example Separate ancestor container queries from viewport queries
 * ```tsx
 * import * as React from 'react';
 * import { useSx, container, width, flex, responsive, viewport, gap } from '@apvee/spfx-react-toolkit/styles';
 * function ResponsiveSummary() {
 *   const sx = useSx();
 *   return <section className={sx(container.inlineSize, width.full)}>
 *     <div className={sx(flex.column, responsive.medium(flex.row), viewport.large(gap.large), gap.medium)}>
 *       <span>First</span><span>Second</span>
 *     </div>
 *   </section>;
 * }
 * ```
 */
export function useSx(options?: SxOptions): SxFunction {
  const { renderer, dir } = useSxStyleEnvironment(options);
  return useMemo(() => (...inputs) => resolveSxInputs({ renderer, dir }, inputs), [renderer, dir]);
}
