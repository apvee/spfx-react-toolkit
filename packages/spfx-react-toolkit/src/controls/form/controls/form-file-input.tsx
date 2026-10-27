import * as React from 'react';
import { mergeSettingsFieldIds } from '../../settings-field-context.internal';
import { Button, Field } from '@fluentui/react-components';
import { useMergedRefs } from '@fluentui/react-utilities';
import { Attach20Regular, Delete20Regular, Document20Regular } from '@fluentui/react-icons';
import type { FieldValues } from 'react-hook-form';
import type { FormFileInputProps } from '../form-control.types';
import { useFieldAdapter } from '../use-field-adapter.internal';
import { useFormRuntimeContext } from '../form-context.internal';
import { registerFormFileInputOwner } from '../form-methods.internal';
import { useSx } from '../../../hooks/useSx';
import * as flex from '../../../helpers/styles/flex';
import * as gap from '../../../helpers/styles/gap';
import * as alignItems from '../../../helpers/styles/align-items';
import { useFormControlsStyles } from './form-controls.styles.internal';

const formatFileNames = (files: FileList): string => Array.from(files, file => file.name).join(', ');

/** Local native file chooser; RHF receives the actual FileList before consumer callbacks.
 * Empty native selection/clear writes null, following HTMLInputElement.files interoperability.
 * No files are uploaded or reconstructed from form defaults; the public ref targets the input.
 * @example `<Order.FileInput name="attachments" label="Attachments" multiple />`
 */
export const FormFileInput = React.forwardRef<HTMLInputElement, FormFileInputProps<FieldValues>>(function FormFileInput(
  { browseLabel: suppliedBrowseLabel, className, clearLabel: suppliedClearLabel, disabled, fieldProps, formatFiles = formatFileNames, hint, id, label, name, onBlur, onChange, orientation, required, rules, ...props }, ref
) {
  const { methods, strings } = useFormRuntimeContext();
  const browseLabel = suppliedBrowseLabel ?? strings.fileInput.browseLabel;
  const clearLabel = suppliedClearLabel ?? strings.fileInput.clearLabel;
  // The browser native file API explicitly uses null for an empty selection.
  // eslint-disable-next-line @rushstack/no-new-null
  const binding = useFieldAdapter<FieldValues, FileList | null>({ describedBy: props['aria-describedby'], className, disabled, fieldProps, hint, id, label, name, orientation, ref, required, rules });
  const inputRef = React.useRef<HTMLInputElement>();
  // React callback refs explicitly receive null during native element detachment.
  // eslint-disable-next-line @rushstack/no-new-null
  const captureInput = React.useCallback((input: HTMLInputElement | null) => { inputRef.current = input ?? undefined; }, []);
  const mergedRef = useMergedRefs<HTMLInputElement>(binding.controller.ref, captureInput);
  const styles = useFormControlsStyles();
  const sx = useSx();
  const isDisabled = binding.controlProps.disabled === true;
  React.useLayoutEffect(() => registerFormFileInputOwner(methods, binding.resetOwner, {
    read: () => inputRef.current?.files ?? undefined,
    clear: () => { if (inputRef.current) inputRef.current.value = ''; },
  }), [methods, binding.resetOwner, binding.name, binding.controller.value]);
  const files = binding.controller.value;
  return <Field {...binding.fieldProps}>{fieldControl => <div className={sx(flex.row, flex.wrap, alignItems.center, gap.small, styles.fileControls)}>
    <input {...props} {...fieldControl} {...binding.controlProps}
      aria-describedby={mergeSettingsFieldIds(fieldControl['aria-describedby'], binding.controlProps['aria-describedby'], props['aria-describedby'])}
      aria-invalid={fieldControl['aria-invalid'] ?? props['aria-invalid']}
      className={styles.fileHidden} ref={mergedRef} name={binding.name} tabIndex={-1} type="file"
      onBlur={event => { binding.controller.onBlur(); onBlur?.(event); }}
      onChange={event => {
        if (isDisabled) return;
        const selected = event.currentTarget.files;
        // Native empty selection has the explicit null contract, including cancel/clear.
        // eslint-disable-next-line @rushstack/no-new-null
        binding.controller.onChange(selected && selected.length ? selected : null);
        onChange?.(event);
      }} />
    <Button disabled={isDisabled} icon={<Attach20Regular aria-hidden="true" />} type="button" onClick={() => { if (!isDisabled) inputRef.current?.click(); }}>{browseLabel}</Button>
    <div className={sx(flex.row, alignItems.center, gap.extraSmall, styles.fileSelection)}>
      {files && files.length > 0 && <Document20Regular aria-hidden="true" />}
      <span aria-live="polite" className={styles.fileSelectionText}>{files && files.length > 0 ? formatFiles(files) : undefined}</span>
      {files && files.length > 0 && <Button appearance="subtle" disabled={isDisabled} icon={<Delete20Regular aria-hidden="true" />} size="small" type="button" onClick={() => {
        if (isDisabled) return;
        if (inputRef.current) inputRef.current.value = '';
        // Native file clear uses null; no synthetic FileList is assigned to the browser.
        // eslint-disable-next-line @rushstack/no-new-null
        binding.controller.onChange(null);
      }}><span className={styles.fileHidden}>{clearLabel}</span></Button>}
    </div>
  </div>}</Field>;
});
FormFileInput.displayName = 'Form.FileInput';
