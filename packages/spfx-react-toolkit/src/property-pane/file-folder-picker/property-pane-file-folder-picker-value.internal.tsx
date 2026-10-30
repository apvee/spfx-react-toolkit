import * as React from 'react';
import { useStableCallback } from '../../hooks/useStableCallback';
import { FileFolderPicker } from '../../controls/file-folder-picker/file-folder-picker';
import type { FileFolderPickerChangeData, FileFolderPickerProps, FileFolderPickerValue } from '../../controls/file-folder-picker/file-folder-picker.types';
import { validateFileFolderPickerValues } from '../../controls/file-folder-picker/file-folder-picker-values.internal';
import { PickerAuthoritativeEpochContext } from '../../controls/pickers/picker-runtime-context.internal';

type DistributiveOmit<T, Key extends PropertyKey> = T extends unknown ? Omit<T, Key> : never;

interface ValueBridgeProps {
  readonly controlProps: DistributiveOmit<FileFolderPickerProps, 'value' | 'defaultValue' | 'onChange'>;
  readonly value: readonly FileFolderPickerValue[];
  readonly valueRevision?: string | number;
  readonly targetProperty: string;
  readonly onChange?: (targetProperty: string, value: readonly FileFolderPickerValue[], isValidEntry: boolean) => void;
}
interface ValueState {
  readonly signature?: string;
  readonly revision?: string | number;
  readonly values: readonly FileFolderPickerValue[];
  readonly draft?: readonly FileFolderPickerValue[];
}

/** Drive references and remote pick actions cannot use the legacy PickerValue/renderer-alias bridge.
 * Reuse the native mount/host/lease, adapting only direct-array draft ownership here. No render/hydration writes.
 */
export function PropertyPaneFileFolderPickerValueBridge(props: ValueBridgeProps): React.ReactElement {
  let values: readonly FileFolderPickerValue[] = [], invalid: string | undefined;
  try { values = validateFileFolderPickerValues(props.value); }
  catch (error) { invalid = error instanceof Error ? error.message : 'Invalid FileFolderPicker JSON value.'; }
  const signature = invalid === undefined ? JSON.stringify(values) : undefined;
  const [state, setState] = React.useState<ValueState>(() => ({ signature, revision: props.valueRevision, values }));
  const replace = invalid === undefined && (signature !== state.signature || props.valueRevision !== state.revision);
  const selection = replace ? values : state.draft ?? state.values;
  const [failure, setFailure] = React.useState<string>();
  React.useLayoutEffect(() => {
    if (replace) {
      setState({ signature, revision: props.valueRevision, values });
      setFailure(undefined);
    }
  }, [replace, signature, props.valueRevision]);
  const mounted = React.useRef(true);
  React.useEffect(() => () => { mounted.current = false; }, []);
  const onChange = useStableCallback((data: FileFolderPickerChangeData): void => {
    if (!mounted.current || invalid !== undefined || props.controlProps.disabled || props.onChange === undefined) return;
    try {
      if (!['pick', 'remove', 'clear'].includes(data.reason)) throw new Error('Value changes require an explicit file/folder picker action.');
      const next = validateFileFolderPickerValues(data.values);
      // Native callbacks may update/mutate their property bag; keep the pending draft independently owned.
      const emitted = validateFileFolderPickerValues(next);
      setState(previous => ({ ...previous, draft: next }));
      setFailure(undefined);
      props.onChange(props.targetProperty, emitted, true);
    } catch (error) { setFailure(error instanceof Error ? error.message : 'Invalid file/folder picker change.'); }
  });
  const missingCallback = props.onChange === undefined ? 'La callback nativa del Property Pane non è disponibile.' : undefined;
  return <PickerAuthoritativeEpochContext.Provider value={JSON.stringify([signature, props.valueRevision])}>
    <FileFolderPicker {...props.controlProps} value={selection} defaultValue={undefined} onChange={onChange}
      disabled={props.controlProps.disabled || invalid !== undefined || props.onChange === undefined}
      validationMessage={invalid ?? failure ?? missingCallback ?? props.controlProps.validationMessage} />
  </PickerAuthoritativeEpochContext.Provider>;
}
