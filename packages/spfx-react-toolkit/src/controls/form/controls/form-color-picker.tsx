import * as React from 'react';
import { Field, ColorPicker, colorAreaClassNames, colorSliderClassNames, alphaSliderClassNames } from '@fluentui/react-components';
import type { FieldValues } from 'react-hook-form';
import type { FormColorPickerProps } from '../form-control.types';
import type { FormHsvColor } from '../form-value.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { useCompositeFieldFocus } from '../use-composite-field-focus.internal';
import { useSx } from '../../../hooks/useSx';
import * as borderWidth from '../../../helpers/styles/border-width';
import * as margin from '../../../helpers/styles/margin';
import * as padding from '../../../helpers/styles/padding';
import * as width from '../../../helpers/styles/width';
import * as maxWidth from '../../../helpers/styles/max-width';
import * as minWidth from '../../../helpers/styles/min-width';
import { useFormControlsStyles } from './form-controls.styles.internal';

function copyColor(color: FormHsvColor): FormHsvColor {
  return { h: color.h, s: color.s, v: color.v, ...(color.a === undefined ? {} : { a: color.a }) };
}
/** Fluent ColorPicker binding with a public root div ref and native descendant error focus.
 * The color area fits its available width with a square aspect ratio, capped at 300px.
 * The picker defaults to a named group; consumer channel names remain unchanged.
 * Native disabled fieldset semantics and capture guards prevent disabled edits.
 * RHF receives changes and blur before consumer callbacks; Field owns accessible relationships.
 * @example `<Order.ColorPicker name="accent" label="Choice" />`
 */
export const FormColorPicker: React.ForwardRefExoticComponent<FormColorPickerProps<FieldValues> & React.RefAttributes<HTMLDivElement>> = React.forwardRef<HTMLDivElement, FormColorPickerProps<FieldValues>>(function FormColorPicker(
  { className, disabled, fieldProps, hint, id, label, name, onBlur, onColorChange, orientation, required, rules, onClickCapture, onKeyDownCapture, ...props }, ref
) {
  const binding = useFieldAdapter<FieldValues, FormHsvColor>({ className, disabled, fieldProps, hint, id, label, name, orientation, required, rules });
  const compositeRef = useCompositeFieldFocus(ref, binding.focusRef, `input.${colorAreaClassNames.inputX},input.${colorAreaClassNames.inputY},input.${colorSliderClassNames.input},input.${alphaSliderClassNames.input}`, () => true);
  const sx = useSx(); const styles = useFormControlsStyles();
  const isDisabled = binding.controlProps.disabled === true;
  // Fluent filters pointer capture props on its root; the native wrapper owns that guard.
  return <Field {...binding.fieldProps}>{fieldControl => <fieldset disabled={isDisabled}
    onPointerDownCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } }}
    className={sx(borderWidth.none, margin.none, padding.none, styles.fieldset)}>
    <ColorPicker {...props} {...fieldControl} className={sx(styles.colorPicker, width.full, maxWidth.px(300), minWidth.zero)} id={binding.controlProps.id} ref={compositeRef} role={props.role ?? 'group'}
      aria-disabled={isDisabled || props['aria-disabled']}
      aria-describedby={[fieldControl['aria-describedby'], props['aria-describedby']].filter(Boolean).join(' ') || undefined}
       color={copyColor(binding.controller.value ?? { h: 0, s: 0, v: 0 })}
      onClickCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } onClickCapture?.(event); }}
      onKeyDownCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } onKeyDownCapture?.(event); }}
      onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
      onColorChange={(event, data) => { if (isDisabled) return; binding.controller.onChange(copyColor(data.color)); onColorChange?.(event, data); }} />
  </fieldset>}</Field>;
});
FormColorPicker.displayName = 'Form.ColorPicker';
