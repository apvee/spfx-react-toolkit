import * as React from 'react';
import { Button } from '@fluentui/react-components';
import type { FormResetProps } from './form.types';
import { useFormRuntimeContext } from './form-context.internal';
/** Native reset button; Root invokes onReset first and honors preventDefault cancellation.
 * @example `<Order.Reset>Discard edits</Order.Reset>`
 */
export const FormReset = React.forwardRef<HTMLButtonElement, FormResetProps>(function FormReset({ disabled, ...props }, ref) {
    const runtime = useFormRuntimeContext();
    return <Button {...props} ref={ref} type="reset" disabled={runtime.disabled || runtime.methods.formState.isSubmitting || disabled === true}/>;
});
FormReset.displayName = 'Form.Reset';
