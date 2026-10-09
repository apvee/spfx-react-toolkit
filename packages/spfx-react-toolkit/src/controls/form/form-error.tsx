import * as React from 'react';
import { useSx } from '../../hooks/useSx';
import * as typography from '../../helpers/styles/typography';
import * as foreground from '../../helpers/styles/foreground';
import { get } from 'react-hook-form';
import type { FieldValues } from 'react-hook-form';
import type { FormErrorProps } from './form.types';
import { useFormRuntimeContext } from './form-context.internal';
import { useResolvedFieldPath } from './field-path.internal';
import { getFormErrorMessages, getGlobalFormErrorMessages, renderFormErrorMessages } from './error-messages.internal';
/** Native alert for a field/array or the form's global errors; duplicate messages render once.
 * @example `<Order.Error />` or `<Order.Error name="reference" />`
 */
export const FormError = React.forwardRef<HTMLDivElement, FormErrorProps<FieldValues>>(function FormError({ name, className, ...props }, ref) {
    const sx = useSx();
    const { methods } = useFormRuntimeContext();
    const resolvedName = useResolvedFieldPath(name ?? '');
    const errors = methods.formState.errors;
    const error: unknown = name === undefined ? undefined : get(errors, resolvedName);
    const arrayRoot: unknown = error && typeof error === 'object' && 'root' in error ? error.root : undefined;
    const messages = name === undefined ? getGlobalFormErrorMessages(errors.root, errors.__formValidation) : getFormErrorMessages(error, arrayRoot);
    if (!messages.length)
        return <></>;
    return <div {...props} className={sx(typography.body1, foreground.primary, className)} ref={ref} role="alert">{renderFormErrorMessages(messages)}</div>;
});
FormError.displayName = 'Form.Error';
