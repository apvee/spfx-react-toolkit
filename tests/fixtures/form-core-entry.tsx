import * as React from 'react';
import { Field } from '@fluentui/react-components';
import type { RegisterOptions } from 'react-hook-form';
import { useFieldAdapter } from '../../packages/spfx-react-toolkit/src/controls/form/use-field-adapter.internal';
export { Form, createForm } from '../../packages/spfx-react-toolkit/src/controls/form';
export { registerFormResetArrayPath, synchronizeFormResetArrayOperation, commitFormResetArrayStructure } from '../../packages/spfx-react-toolkit/src/controls/form/form-methods.internal';
/** Real private binding in the same bundle as public suite; simulates incomplete native draft. */
export function DraftControl({ rules, name = 'title' }: { rules?: RegisterOptions<Record<string, string>>; name?: string }): React.ReactElement {
  const [draft, setDraft] = React.useState('');
  const binding = useFieldAdapter<Record<string, string>, string>({ name, rules, validateInput: () => draft === 'incomplete' ? 'Draft incomplete' : true });
  return <Field {...binding.fieldProps}><input {...binding.controlProps} ref={binding.controller.ref} value={draft} onChange={event => { setDraft(event.target.value); }} /></Field>;
}

/** Draft presentation used to observe public/reset coordinator timing without a second values store. */
export function ResetDraftControl(): React.ReactElement {
  const [draft, setDraft] = React.useState('');
  const binding = useFieldAdapter<Record<string, string>, string>({ name: 'title' });
  React.useEffect(() => { if (binding.resetVersion && !binding.resetPreservesValues) setDraft(''); }, [binding.resetVersion, binding.resetPreservesValues]);
  return <><input data-draft="true" ref={binding.controller.ref} value={draft} onChange={event => setDraft(event.target.value)} /><output data-reset="true">{`${binding.resetVersion}:${binding.resetSource ?? ''}`}</output></>;
}
