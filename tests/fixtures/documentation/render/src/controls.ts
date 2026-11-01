/** Source-owned stage introduction.
 * @packageDocumentation
 * @docPage controls/stage
 */
/** Shared opening state. */
export interface StageBase {
  /** Whether the surface is open. Omitted means closed.
   * @defaultValue false
   */
  open?: boolean;
}
/** Stage presentation constraints. */
export type StageProps = StageBase & (
  { presentation: 'dialog'; trigger?: never } |
  { presentation: 'popover'; trigger: string }
);
/** Opens a stage.
 * @useWhen You need a controlled surface.
 * @avoidWhen A persistent section is enough.
 * @remarks
 * ### Requirements
 * Popover requires a trigger.
 * ### Limitations
 * The caller owns open state.
 * @param props - Stage inputs.
 * @returns The requested presentation.
 * @example Basic
 * ```ts
 * import { Stage } from '@fixture/render';
 * Stage({ presentation: 'popover', trigger: 'Open' });
 * ```
 */
export function Stage(props: StageProps): string { return props.presentation; }
/** Observes invalid submission. */
export interface FormRootProps {
  /** Promise results are observed; rejected Promises reach onError. Omitted means no observer. */
  onInvalid?: (reason: string) => void | Promise<void>;
}
/** Inherits the open state. */
export interface Derived extends StageBase {
  /** Label content. */
  label: string;
}
/** A source-owned scalar. */
export const answer = 42;
/** A readonly source shape. */
export interface ReadonlyBase { readonly size: number; label: string }
/** Inherited readonly behavior. */
export interface ReadonlyChild extends ReadonlyBase { extra: boolean }
/** Added mapped readonly behavior. */
export type Frozen = { readonly [K in keyof ReadonlyBase]: ReadonlyBase[K] };
/** Removed mapped readonly behavior. */
export type Mutable = { -readonly [K in keyof ReadonlyBase]: ReadonlyBase[K] };
/** A second callable on the same route.
 * @param value - Value to preserve.
 * @returns The unchanged value.
 */
export const action = <T extends string = string>(value: T): T => value;
/** Public alias contract. */
export { action as perform };
