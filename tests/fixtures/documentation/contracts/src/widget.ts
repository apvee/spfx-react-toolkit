/** A widget part. */
function Part(): void {}
/** Compound widget.
 * @useWhen You need grouped widget parts.
 * @example Widget
 * ```ts
 * import { widget } from '@fixture/contracts';
 * widget.Part();
 * ```
 */
export const widget = {
  /** Invokes the widget part. */
  Part
};
