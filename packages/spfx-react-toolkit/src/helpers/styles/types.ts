/**
 * Opaque immutable style descriptors and ordered class-composition contracts.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @remarks
 * Descriptor construction/import accesses no DOM and inserts no CSS. Materialize classes with useSx under the host Griffel renderer; token values need scoped Fluent theme variables. States accept only base descriptors; queries accept base/state descriptors. Nested states, nested queries and class strings within scopes are rejected; place class strings and ignored conditional values only in the outer composer.
 */
/** Internal brand; consumers obtain descriptors from toolkit factories. */
export const sxDescriptorBrand: unique symbol = Symbol('apvee.sx');

/**
 * Opaque immutable declaration or recipe, usable directly by the style composer.
 * @example
 * ```ts
 * import { SxBaseDescriptor, width } from '@apvee/spfx-react-toolkit/styles';
 * const input: SxBaseDescriptor = width.full;
 * ```
 */
export interface SxBaseDescriptor {
  /** Internal discriminant preventing raw CSS objects from satisfying the contract. */
  readonly [sxDescriptorBrand]: 'base';
}
/**
 * Opaque CSS state containing base descriptors; states cannot be nested.
 * @example
 * ```ts
 * import { SxStateDescriptor, hover, foreground } from '@apvee/spfx-react-toolkit/styles';
 * const state: SxStateDescriptor = hover(foreground.subtle);
 * ```
 */
export interface SxStateDescriptor {
  /** Internal state descriptor brand. */
  readonly [sxDescriptorBrand]: 'state';
}
/**
 * Opaque container or viewport query containing base or state descriptors.
 * @example
 * ```ts
 * import { SxResponsiveDescriptor, responsive, hover, foreground } from '@apvee/spfx-react-toolkit/styles';
 * const query: SxResponsiveDescriptor = responsive.medium(hover(foreground.subtle));
 * ```
 */
export interface SxResponsiveDescriptor {
  /** Internal responsive descriptor brand. */
  readonly [sxDescriptorBrand]: 'responsive';
}
/**
 * Any supported immutable style descriptor.
 * @example
 * ```ts
 * import { SxDescriptor, hover, foreground } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor: SxDescriptor = hover(foreground.subtle);
 * ```
 */
export type SxDescriptor = SxBaseDescriptor | SxStateDescriptor | SxResponsiveDescriptor;
/**
 * Composer argument: descriptors, compiled classes, or ignored conditional values.
 * @example
 * ```ts
 * import { SxInput, width } from '@apvee/spfx-react-toolkit/styles';
 * const input: SxInput = undefined;
 * const inputs: readonly SxInput[] = [width.full, input];
 * ```
 */
// eslint-disable-next-line @rushstack/no-new-null -- The conditional composer contract explicitly accepts null.
export type SxInput = SxDescriptor | string | false | null | undefined;
/**
 * Ordered style composer; returns a className string, empty when no inputs apply.
 * @param inputs - Ordered descriptors and existing class names; last declaration wins within a scope.
 * @returns Composed CSS class names.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SxFunction, useSx, width, foreground } from '@apvee/spfx-react-toolkit/styles';
 * export function Surface({ subtle }: { subtle: boolean }): React.ReactElement {
 *   const sx: SxFunction = useSx();
 *   return <div className={sx(width.full, subtle && foreground.subtle)}>Content</div>;
 * }
 * ```
 */
export type SxFunction = (...inputs: readonly SxInput[]) => string;
/**
 * Options describing the effective direction of a style environment.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SxOptions, useSx, width } from '@apvee/spfx-react-toolkit/styles';
 * export function Surface(): React.ReactElement {
 *   const options: SxOptions = { dir: 'rtl' };
 *   const sx = useSx(options);
 *   return <div dir="rtl" className={sx(width.full)}>Content</div>;
 * }
 * ```
 */
export interface SxOptions {
  /** Effective direction override; keep consistent with the host and external Griffel classes.
   * When omitted, uses Fluent direction context; without context, defaults to LTR. */
  readonly dir?: 'ltr' | 'rtl';
}
