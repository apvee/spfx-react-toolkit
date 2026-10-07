/** Internal brand; consumers obtain descriptors from toolkit factories. */
export const sxDescriptorBrand: unique symbol = Symbol('apvee.sx');

/** Opaque immutable declaration or recipe, usable directly by the style composer.
 * @example const input: SxBaseDescriptor = width.full;
 */
export interface SxBaseDescriptor {
  /** Internal discriminant preventing raw CSS objects from satisfying the contract. */
  readonly [sxDescriptorBrand]: 'base';
}
/** Opaque CSS state containing base descriptors; states cannot be nested.
 * @example hover(foreground.subtle)
 */
export interface SxStateDescriptor {
  /** Internal state descriptor brand. */
  readonly [sxDescriptorBrand]: 'state';
}
/** Opaque container or viewport query containing base or state descriptors.
 * @example responsive.medium(hover(foreground.subtle))
 */
export interface SxResponsiveDescriptor {
  /** Internal responsive descriptor brand. */
  readonly [sxDescriptorBrand]: 'responsive';
}
/** Any supported immutable style descriptor. */
export type SxDescriptor = SxBaseDescriptor | SxStateDescriptor | SxResponsiveDescriptor;
/** Composer argument: descriptors, compiled classes, or ignored conditional values. */
// eslint-disable-next-line @rushstack/no-new-null -- The conditional composer contract explicitly accepts null.
export type SxInput = SxDescriptor | string | false | null | undefined;
/** Ordered style composer; returns a className string, empty when no inputs apply.
 * @param inputs Ordered descriptors and existing class names; last declaration wins within a scope.
 * @returns Composed CSS class names.
 * @example sx(width.full, condition && foreground.subtle)
 */
export type SxFunction = (...inputs: readonly SxInput[]) => string;
/** Options describing the effective direction of a style environment. */
export interface SxOptions {
  /** Effective direction override; keep consistent with the host and external Griffel classes.
   * When omitted, uses Fluent direction context; without context, defaults to LTR. */
  readonly dir?: 'ltr' | 'rtl';
}
