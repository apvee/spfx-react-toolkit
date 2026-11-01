/**
 * Navigation helpers.
 * @packageDocumentation
 * @docPage controls/pagination
 */
/**
 * Produces a page number.
 * @useWhen You own results and need controlled pages.
 * @avoidWhen You need server fetching; fetch before invoking this helper.
 * @param value - Requested page number.
 * @returns The requested page.
 * @remarks
 * ### Requirements
 * No provider is required.
 * @example Basic
 * ```ts
 * import { paginate } from '@fixture/contracts';
 * const page = paginate(1);
 * // @madeUp is code, not a documentation tag.
 * ```
 */
export function paginate(value: number): number { return value; }
