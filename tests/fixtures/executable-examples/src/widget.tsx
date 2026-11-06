import * as React from 'react';

/**
 * Displays a widget.
 * @docPage controls/widget
 * @scenario Basic view
 * ```tsx
 * import * as React from 'react';
 * import { Widget } from '@apvee/spfx-react-toolkit';
 * function MisleadingFirst(): string { return 'ordinary helper'; }
 * export function Demo(): React.ReactElement { return <Widget />; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A widget is visible.' };
 * ```
 * @example Ordinary component
 * ```tsx
 * import * as React from 'react';
 * export default function Example(): React.ReactElement { return <div />; }
 * ```
 */
export function Widget(): React.ReactElement { return <div>Widget</div>; }

/**
 * Optional widget configuration.
 * @example Type usage
 * ```ts
 * import type { WidgetOptions } from '@apvee/spfx-react-toolkit';
 * const options: WidgetOptions = {};
 * ```
 */
export interface WidgetOptions { enabled?: boolean; }
