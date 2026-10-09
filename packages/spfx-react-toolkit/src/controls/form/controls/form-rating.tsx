import * as React from 'react';
import { Field, Rating, ratingItemClassNames } from '@fluentui/react-components';
import type { FieldValues } from 'react-hook-form';
import type { FormRatingProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { useCompositeFieldFocus } from '../use-composite-field-focus.internal';
import { useSx } from '../../../hooks/useSx';
import * as borderWidth from '../../../helpers/styles/border-width';
import * as margin from '../../../helpers/styles/margin';
import * as padding from '../../../helpers/styles/padding';
import { useFormControlsStyles } from './form-controls.styles.internal';

/** Fluent Rating binding with a public root div ref and native descendant error focus.
 * Defaults to native marigold stars; an explicit color overrides that presentation.
 * Native disabled fieldset semantics and capture guards prevent disabled edits.
 * RHF receives changes and blur before consumer callbacks; Field owns accessible relationships.
 * @example `<Order.Rating name="value" label="Choice" />`
 */
export const FormRating: React.ForwardRefExoticComponent<FormRatingProps<FieldValues> & React.RefAttributes<HTMLDivElement>> = React.forwardRef<HTMLDivElement, FormRatingProps<FieldValues>>(function FormRating(
  { className, color = 'marigold', disabled, fieldProps, hint, id, label, name, onBlur, onChange, orientation, required, rules, onClickCapture, onKeyDownCapture, ...props }, ref
) {
  const binding = useFieldAdapter<FieldValues, number>({ className, disabled, fieldProps, hint, id, label, name, orientation, required, rules });
  const compositeRef = useCompositeFieldFocus(ref, binding.focusRef, `input.${ratingItemClassNames.fullValueInput}[type="radio"],input.${ratingItemClassNames.halfValueInput}[type="radio"]`, element => element.name === binding.name);
  const sx = useSx(); const styles = useFormControlsStyles();
  const isDisabled = binding.controlProps.disabled === true;
  // Fluent filters pointer capture props on its root; the native wrapper owns that guard.
  return <Field {...binding.fieldProps}>{fieldControl => <fieldset disabled={isDisabled}
    onPointerDownCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } }}
    className={sx(borderWidth.none, margin.none, padding.none, styles.fieldset)}>
    <Rating {...props} {...fieldControl} color={color} id={binding.controlProps.id} ref={compositeRef}
      aria-disabled={isDisabled || props['aria-disabled']}
      aria-describedby={[fieldControl['aria-describedby'], props['aria-describedby']].filter(Boolean).join(' ') || undefined}
      name={binding.name} value={binding.controller.value ?? 0}
      onClickCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } onClickCapture?.(event); }}
      onKeyDownCapture={event => { if (isDisabled) { event.preventDefault(); event.stopPropagation(); } onKeyDownCapture?.(event); }}
      onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
      onChange={(event, data) => { if (isDisabled) return; binding.controller.onChange(data.value); onChange?.(event, data); }} />
  </fieldset>}</Field>;
});
FormRating.displayName = 'Form.Rating';
