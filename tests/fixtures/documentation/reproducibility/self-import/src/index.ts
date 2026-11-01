import type { NativeOptions } from '@fixture/cold/native';
/**
 * Read the dedicated native contract through a clean self import.
 * @docPage helpers/cold
 * @param options - Required native options.
 * @returns The selected name.
 * @example Clean imports
 * ```ts
 * import { control } from '@fixture/cold';
 * import type { NativeOptions } from '@fixture/cold/native';
 * const options: NativeOptions = { name: 'cold' };
 * const result: string = control(options);
 * ```
 */
export function control(options: NativeOptions): string { return options.name; }
