import * as React from 'react';
import { Card, Field, useId } from '@fluentui/react-components';
import type { FieldValues } from 'react-hook-form';
import type { FormObjectPath, FormObjectSectionProps } from '../form-composition.types';
import { FieldPathScope, useFieldScopeDisabled, useResolvedFieldPath } from '../field-path.internal';
import { useFormRuntimeContext } from '../form-context.internal';
import { FormObject } from './form-object';
import { assertSectionName, sectionDescriptionIds, useFormSectionsStyles } from './form-sections.styles.internal';
import { GroupText, hasGroupContent, hasGroupDescriptionContent } from '../../group-header.internal';
import { renderContentSlot } from '../../content-slot.internal';
import { useSx } from '../../../hooks/useSx';
import * as flex from '../../../helpers/styles/flex';
import * as gap from '../../../helpers/styles/gap';
import * as minWidth from '../../../helpers/styles/min-width';
import * as marginBlockStart from '../../../helpers/styles/margin-block-start';

/** Named native fieldset around the existing record scope, with Fluent description/errors.
 * Direct fields use medium spacing; descriptions/errors use small spacing.
 * Disabled scopes reach RHF adapters; native disabling alone is insufficient for submission omission.
 * @param props Relative path, accessible group name, native attributes and nested field renderer.
 * @returns Accessible record section; caller children own custom layout.
 * @example `<Order.ObjectSection name="customer" title="Customer">{({ Fields }) => <Fields.Input name="name" />}</Order.ObjectSection>`
 */
export function FormObjectSection<TInput extends FieldValues, TPath extends FormObjectPath<TInput>>({ name, children, title, description, disabled: ownDisabled, appearance = 'filled', className, 'aria-describedby': describedBy, ...native }: FormObjectSectionProps<TInput, TPath>): React.ReactElement {
    assertSectionName(title, native['aria-label'], native['aria-labelledby']);
    const runtime = useFormRuntimeContext();
    const scopeDisabled = useFieldScopeDisabled();
    const disabled = runtime.disabled || scopeDisabled || ownDisabled === true;
    const prefix = useResolvedFieldPath('');
    const sx = useSx(), styles = useFormSectionsStyles();
    const descriptionId = useId('form-object-description'), errorId = useId('form-object-error');
    const hasDescription = hasGroupDescriptionContent(description);
    return <FieldPathScope prefix={prefix} disabled={disabled}><FormObject<TInput, TPath> name={name}>{context =>
        <fieldset {...native} disabled={disabled} aria-invalid={context.errors.length > 0 || native['aria-invalid']} aria-describedby={sectionDescriptionIds(describedBy, hasDescription ? descriptionId : undefined, context.errors.length ? errorId : undefined)} className={sx(styles.group, minWidth.zero, className)}>
            {hasGroupContent(title) && <legend className={styles.legend}><GroupText>{title}</GroupText></legend>}
            <Card appearance={appearance} focusMode="off" size="medium">
                <Field className={sx(gap.small)} hint={hasGroupContent(description) ? { id: descriptionId, className: sx(marginBlockStart.none), children: renderContentSlot(description, 'div', sx, { shorthand: 'preserve' }) } : undefined} validationState={context.errors.length ? 'error' : 'none'} validationMessage={context.errors.length ? { id: errorId, className: sx(marginBlockStart.none), children: context.errors.join('; ') } : undefined}>
                    <div data-form-object-content="" className={sx(flex.column, gap.medium, minWidth.zero)}>{children(context)}</div>
                </Field>
            </Card>
        </fieldset>
    }</FormObject></FieldPathScope>;
}
