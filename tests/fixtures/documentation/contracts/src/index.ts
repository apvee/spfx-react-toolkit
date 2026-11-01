/**
 * Aggregate exports do not own implementation routes.
 * @packageDocumentation
 * @docPage aggregate
 */
export { paginate, paginate as pages } from './pagination';
export type { PagingOptions } from './types';
/**
 * Alias-specific navigation binding.
 * @useWhen A compact public alias is convenient.
 * @docPage controls/pagination
 * @example Alias
 * ```ts
 * import { navigate } from '@fixture/contracts';
 * navigate(1);
 * ```
 */
export { paginate as navigate } from './pagination';
export { Original, Derived, Customized } from './inheritance';
export { widget } from './widget';
export { identity } from './identity';
