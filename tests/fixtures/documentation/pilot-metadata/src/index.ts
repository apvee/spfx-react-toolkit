import { Compose, ModelHandler, ModelSelector } from './contracts';
import { upstream } from '../peer';
export { Compose, choose, Selector, SeparateSelector, ExtendedChoice, configure, ModelHandler, ModelSelector } from './contracts';
export * as measures from './measures';
/** A local binding override.
 * @useWhen Select the alternate local context.
 * @example
 * ```ts
 * import { alternate } from '@fixture/pilot';
 * const value = alternate.unit(1);
 * ```
 */
export * as alternate from './measures';
/** Constructs a value. See {@link Compose} and {@link assemble}.
 * @docPage shared
 * @useWhen Build a stable value.
 * @example
 * ```ts
 * import { compose } from '@fixture/pilot';
 * const value = compose();
 * ```
 */
export declare const compose: Compose;
export { compose as assemble };
/** Toolkit alias contract with local timing requirements.
 * @docPage shared
 * @useWhen Keep the local callable contract.
 * @typeParam T - Local input model.
 * @param value - Required local value.
 * @returns Local result.
 * @example
 * ```ts
 * import { externalAlias } from '@fixture/pilot';
 * const value = externalAlias(1);
 * ```
 */
export const externalAlias = upstream;

/** Processes a number through a closed model handler.
 * @docPage shared
 * @useWhen Process number values.
 * @example
 * ```ts
 * import { numberHandler } from '@fixture/pilot';
 * const value: number = numberHandler(1);
 * ```
 */
export declare const numberHandler: ModelHandler<number>;
/** Selects keys from a closed record model.
 * @docPage shared
 * @useWhen Select known named record keys.
 * @example
 * ```ts
 * import { nameSelector } from '@fixture/pilot';
 * const value: 'name' = nameSelector('name');
 * ```
 */
export declare const nameSelector: ModelSelector<{ name: string }>;
