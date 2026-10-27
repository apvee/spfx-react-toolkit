import * as React from 'react';
import { mergeSettingsFieldIds } from '../../settings-field-context.internal';
import { Field, SwatchPicker, colorSwatchClassNames, imageSwatchClassNames } from '@fluentui/react-components';
import type { FieldValues } from 'react-hook-form';
import type { FormSwatchPickerProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { useCompositeFieldFocus } from '../use-composite-field-focus.internal';
import { useSx } from '../../../hooks/useSx';
import * as borderWidth from '../../../helpers/styles/border-width';
import * as margin from '../../../helpers/styles/margin';
import * as padding from '../../../helpers/styles/padding';
import { useFormControlsStyles } from './form-controls.styles.internal';

/** Fluent SwatchPicker binding with a public root div ref and native descendant error focus.
 * Native disabled fieldset semantics and capture guards prevent disabled edits.
 * RHF receives changes and blur before consumer callbacks; Field owns accessible relationships.
 * Swatches use Fluent's rounded shape by default; the native shape prop overrides it.
 * @example `<Order.SwatchPicker name="value" label="Choice" />`
 */
export const FormSwatchPicker: React.ForwardRefExoticComponent<FormSwatchPickerProps<FieldValues> & React.RefAttributes<HTMLDivElement>> = React.forwardRef<HTMLDivElement, FormSwatchPickerProps<FieldValues>>(function FormSwatchPicker(
  { className, disabled, fieldProps, hint, id, label, name, onBlur, onSelectionChange, orientation, required, rules, shape = 'rounded', onClickCapture, onKeyDownCapture, ...props }, ref
) {
  const binding = useFieldAdapter<FieldValues, string>({ settingsLabelTarget: 'group', describedBy: props['aria-describedby'], className, disabled, fieldProps, hint, id, label, name, orientation, required, rules });
  const compositeRef = useCompositeFieldFocus<HTMLDivElement, HTMLButtonElement>(ref, binding.focusRef, `button.${colorSwatchClassNames.root},button.${imageSwatchClassNames.root}`, () => true);
  const sx = useSx(); const styles = useFormControlsStyles();
  const isDisabled = binding.controlProps.disabled === true;
  // Fluent filters pointer capture props on its root; the native wrapper owns that guard.
  return <Field {...binding.fieldProps}>{fieldControl => <fieldset disabled={isDisabled}
    onPointerDownCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } }}
    className={sx(borderWidth.none, margin.none, padding.none, styles.fieldset)}>
    <SwatchPicker {...props} {...fieldControl} shape={shape} aria-labelledby={binding.controlProps['aria-labelledby'] ?? fieldControl['aria-labelledby']} id={binding.controlProps.id} ref={compositeRef}
      aria-disabled={isDisabled || props['aria-disabled']}
      aria-describedby={mergeSettingsFieldIds(binding.controlProps['aria-describedby'], props['aria-describedby'])}
       selectedValue={binding.controller.value ?? ''}
      onClickCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } onClickCapture?.(event); }}
      onKeyDownCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } onKeyDownCapture?.(event); }}
      onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
      onSelectionChange={(event, data) => { if (isDisabled) return; binding.controller.onChange(data.selectedValue); onSelectionChange?.(event, data); }} />
  </fieldset>}</Field>;
});
FormSwatchPicker.displayName = 'Form.SwatchPicker';
