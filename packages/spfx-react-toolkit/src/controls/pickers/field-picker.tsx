import * as React from 'react';
import type { FieldPickerProps, FieldPickerOption } from './picker.types';
import { createFieldPickerSource } from './field-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import type { FieldPickerValue } from './picker-values.types';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/** Search list fields or site columns beneath an SPFx provider, with icons based on verified field types.
 * Defaults: Dropdown (historical combobox presentation), multiple selection, hidden/read-only included, Title ascending.
 * Opening the menu loads all matching fields, following native paging. Typed search filters the complete memory-cached catalog locally, case-insensitively; no initial display limit.
 * @param props List selector (omitted for site columns), references and Fluent options.
 * @returns A native Fluent Field and picker; blank clearing emits an empty array.
 * @example <FieldPicker label="Campi" list="Projects" includeReadOnly={false} />
 */
export function FieldPicker(props: FieldPickerProps): React.ReactElement {
  const context = usePickerSourceContext(), source = createFieldPickerSource(props, context);
  const runtime = usePickerRuntime<FieldPickerValue, FieldPickerOption>(props, source, { multiselect: true });
  return <PickerRenderer<FieldPickerValue, FieldPickerOption> props={props} runtime={runtime} defaultPresentation="combobox" defaultMultiselect showBlankOption={props.showBlankOption} />;
}
