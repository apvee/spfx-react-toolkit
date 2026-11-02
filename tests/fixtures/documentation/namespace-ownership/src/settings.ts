/**
 * Settings with an implemented default expression.
 * @packageDocumentation
 * @docPage helpers/styles
 * @example Default expression recipe
 * ```tsx
 * import * as React from 'react';
 * import { settings } from '@fixture/namespace';
 * export function Demo(): React.ReactElement { return <div>{settings.default.actual}</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'The default settings value is visible.' };
 * ```
 */
export default { actual: 1 };
