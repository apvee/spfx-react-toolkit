import { makeStyles, shorthands } from '@griffel/react';
import type * as React from 'react';
import { hasGroupContent } from '../../group-header.internal';

/** Details absent from existing layout/typography descriptors. No renderer is created here. */
export const useFormSectionsStyles = makeStyles({
    group: { ...shorthands.border(0), margin: 0, padding: 0, minInlineSize: 0, '&[hidden]': { display: 'none' } },
    legend: { padding: 0, marginBlockEnd: '8px', maxInlineSize: '100%' },
    list: { listStyleType: 'none', margin: 0, padding: 0 },
    item: { minInlineSize: 0 },
    live: { position: 'absolute', width: '1px', height: '1px', padding: 0, margin: '-1px', overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap', ...shorthands.border(0) },
});
/** Runtime guard also covers untyped/dynamic callers and empty legends. */
export function assertSectionName(title: React.ReactNode, label: string | undefined, labelledBy: string | undefined): void {
    if (!hasGroupContent(typeof title === 'string' ? title.trim() : title) && !label?.trim() && !labelledBy?.trim()) throw new TypeError('Form sections require an accessible name through title, aria-label or aria-labelledby.');
}
/** Adds direct section description/error references to caller-owned references. */
export function sectionDescriptionIds(existing: string | undefined, descriptionId: string | undefined, errorId: string | undefined): string | undefined {
    return [existing, descriptionId, errorId].filter(Boolean).join(' ') || undefined;
}
