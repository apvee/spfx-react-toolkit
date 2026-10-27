import * as React from 'react';
import { Button } from '@fluentui/react-components';
import type { FormSubmitProps } from './form.types';
import { useFormRuntimeContext } from './form-context.internal';
/** Native submit button; RHF pending and Root disabled state cannot be overridden.
 * @example `<Order.Submit>Save</Order.Submit>`
 */
export const FormSubmit = React.forwardRef<HTMLButtonElement, FormSubmitProps>(function FormSubmit({ disabled, ...props }, ref) {
    const runtime = useFormRuntimeContext();
    return <Button {...props} ref={ref} type="submit" disabled={runtime.disabled || runtime.methods.formState.isSubmitting || disabled === true}/>;
});
FormSubmit.displayName = 'Form.Submit';
