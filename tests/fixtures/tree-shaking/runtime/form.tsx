import * as React from 'react';
import { act, Simulate } from 'react-dom/test-utils';
import { Form, createForm } from '__TOOLKIT__';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import { useFormContext } from 'react-hook-form';

const Order = createForm({ '~standard': { version: 1, vendor: 'production-package',
  validate: async value => ({ value: { amount: Number(value.amount), date: value.date } }) } });

export async function run({ ReactDOM, assert }) {
  const container = document.createElement('div'); document.body.appendChild(container);
  const forms = {}, fields = {}, contexts = {}, directions = {}, submissions = [];
  let releaseSave;
  const pendingSave = new Promise(resolve => { releaseSave = resolve; });
  const defaults = { amount: '3', date: new Date(2026, 9, 8) };
  function ReadField({ id }) {
    fields[id] = Order.useField({ name: 'amount' });
    contexts[id] = useFormContext();
    directions[id] = useFluent_unstable().dir;
    return <output data-field={id}>{fields[id].value}</output>;
  }
  function Record({ id }) {
    const methods = Order.useForm({ defaultValues: defaults });
    forms[id] = methods;
    return <Order.Root id={id} methods={methods} onSubmit={async value => { submissions.push({ id, value }); await pendingSave; }}>
      <Order.Input name="amount" label="Amount" id={`${id}-amount`} />
      <Order.DatePicker name="date" label="Date" />
      <ReadField id={id} /><Order.Error />
      <Order.Submit>Save</Order.Submit><Order.Reset>Reset</Order.Reset>
    </Order.Root>;
  }
  try {
    await act(async () => { ReactDOM.render(<FluentProvider theme={webLightTheme} dir="rtl"><Record id="first" /><Record id="second" /></FluentProvider>, container); });
    assert.equal(contexts.first.control, forms.first.control, 'RHF context reaches the same-suite methods');
    assert.equal(contexts.second.control, forms.second.control);
    assert.notEqual(forms.first.control, forms.second.control, 'Form instances own independent state');
    assert.deepEqual(directions, { first: 'rtl', second: 'rtl' });
    assert.equal(fields.first.value, '3');
    const input = container.querySelector('#first-amount');
    await act(async () => { Simulate.change(input, { target: { value: '8' } }); });
    assert.equal(forms.first.getValues('amount'), '8');
    assert.equal(forms.second.getValues('amount'), '3');
    assert.equal(fields.first.value, '8');
    await act(async () => { Simulate.submit(container.querySelector('#first')); Simulate.submit(container.querySelector('#first')); });
    assert.equal(submissions.length, 1, 'Root shares the pending validation/save gate');
    assert.equal(submissions[0].value.amount, 8, 'Validated output transforms the live string');
    assert.equal(forms.first.getValues('amount'), '8', 'Live model preserves input domain');
    assert.ok(submissions[0].value.date instanceof Date);
    assert.notEqual(submissions[0].value.date, forms.first.getValues('date'), 'Submission snapshots dates');
    assert.ok(container.querySelector('#first button[type="submit"]').disabled);
    await act(async () => { releaseSave(); });
    await act(async () => { Simulate.reset(container.querySelector('#first')); });
    assert.equal(forms.first.getValues('amount'), '3');
    assert.equal(container.querySelector('#first-amount').value, '3');
    assert.equal(forms.second.getValues('amount'), '3');
    await act(async () => { forms.first.reset({ amount: '11', date: new Date(2027, 0, 1) }); });
    assert.equal(container.querySelector('#first-amount').value, '11');
    assert.equal(forms.second.getValues('amount'), '3');
    for (const name of ['Root', 'Submit', 'Reset', 'Error', 'Input', 'Textarea', 'SearchBox', 'Checkbox', 'Switch', 'SpinButton', 'Slider', 'Select', 'RadioGroup', 'Dropdown', 'DropdownMulti', 'CheckboxGroup', 'Combobox', 'ComboboxMulti', 'TagPicker', 'Rating', 'ColorPicker', 'SwatchPicker', 'DatePicker', 'TimePicker', 'DateTime', 'FileInput', 'Object', 'Array', 'ObjectSection', 'ArraySection', 'useForm', 'useField']) assert.ok(Form[name], `Missing complete-family member ${name}`);
    return { assertions: ['actual installed suite mounts with shared RHF/Fluent contexts', 'independent form state', 'schema input/output transformation', 'Root pending double-submit gate', 'native reset and explicit record reset', 'Date submission snapshots', 'complete compound family'], submissions: submissions.length, liveAmount: forms.first.getValues('amount'), secondAmount: forms.second.getValues('amount') };
  } finally {
    await act(async () => { ReactDOM.unmountComponentAtNode(container); }); container.remove();
  }
}
