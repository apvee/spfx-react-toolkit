import * as React from 'react';
import { Form as RootForm, createForm as rootCreateForm } from '@apvee/spfx-react-toolkit';
import { Form as DomainForm, createForm as domainCreateForm } from '@apvee/spfx-react-toolkit/controls';
import { Form, createForm } from '@apvee/spfx-react-toolkit/controls/form';
import type { FormMethods, FormDateTimeValue, FormHsvColor, FormTimePickerProps } from '@apvee/spfx-react-toolkit/controls/form';
import type { InputProps } from '@fluentui/react-components';
import type { DatePickerProps } from '@fluentui/react-datepicker-compat';
import type { TimePickerProps } from '@fluentui/react-timepicker-compat';
import type { StandardSchemaV1 } from '@standard-schema/spec';
import type { UseFormReturn } from 'react-hook-form';

export const aliases: typeof Form[] = [RootForm, DomainForm];
export const factoryAliases: typeof createForm[] = [rootCreateForm, domainCreateForm];
const schema: StandardSchemaV1<{ amount: string; rows: { title: string }[] }, { amount: number }> = {
  '~standard': { version: 1, vendor: 'package-consumer', validate: value => ({ value: { amount: Number((value as { amount: string }).amount) } }) }
};
const Order = createForm(schema);
export function TypedConsumer(): React.ReactElement {
  const methods = Order.useForm({ defaultValues: { amount: '3', rows: [{ title: 'One' }] } });
  const live: string = methods.getValues('amount');
  const native: UseFormReturn<{ amount: string; rows: { title: string }[] }, unknown, { amount: number }> = methods;
  const declared: FormMethods<{ amount: string; rows: { title: string }[] }, unknown, { amount: number }> = methods;
  // @ts-expect-error Raw RHF methods do not carry suite ownership.
  const unowned: FormMethods<{ amount: string; rows: { title: string }[] }, unknown, { amount: number }> = native;
  void unowned;
  const input = React.createRef<HTMLInputElement>();
  const onChange: NonNullable<InputProps['onChange']> = (event, data) => {
    const target: HTMLInputElement = event.currentTarget;
    methods.setValue('amount', data.value);
    void target;
  };
  void live; void declared;
  return <Order.Root methods={methods} onSubmit={value => {
    value.amount.toFixed();
    // @ts-expect-error Validation transforms the submission, not live input values.
    value.amount.toUpperCase();
  }}>
    <Order.Input name="amount" ref={input} onChange={onChange} />
    <Order.Array name="rows">{({ items }) => items.map(({ key, Fields }) => <Fields.Input key={key} name="title" />)}</Order.Array>
    <Order.Error  /><Order.Submit>Save</Order.Submit><Order.Reset>Reset</Order.Reset>
  </Order.Root>;
}

type Model = { date: Date | null; time: Date | null; appointment: FormDateTimeValue; color: FormHsvColor; file: FileList | null; flag: boolean; count: number | null; tags: string[]; title: string };
const Fields = createForm<Model>();
export function CompatConsumer(): React.ReactElement {
  const dateRef = React.createRef<HTMLInputElement>();
  const fileRef = React.createRef<HTMLInputElement>();
  const dateCallback: DatePickerProps['onSelectDate'] = selected => { void selected; };
  const nativeTimeProps: TimePickerProps = { onTimeChange: (event, data) => { void event; void data.selectedTime; } };
  const timeCallback: FormTimePickerProps<Model>['onTimeChange'] = (event, data) => { const target: HTMLElement = event.currentTarget; const selected: Date | null = data.selectedTime; void target; void selected; };
  void nativeTimeProps;
  return <Fields.Root defaultValues={{ date: null, time: null, appointment: { date: null, time: null }, file: null, count: null, flag: false, tags: [], title: '', color: { h: 0, s: 1, v: 1 } }} onSubmit={value => { value.title.toUpperCase(); }}>
    <Fields.DatePicker name="date" ref={dateRef} onSelectDate={dateCallback} />
    <Fields.TimePicker name="time" onTimeChange={timeCallback} />
    <Fields.DateTime name="appointment" /><Fields.ColorPicker name="color" />
    <Fields.FileInput name="file" ref={fileRef} /><Fields.SpinButton name="count" />
    <Fields.Checkbox name="flag" /><Fields.TagPicker name="tags" label="Tags" options={[]} />
  </Fields.Root>;
}
// @ts-expect-error Numeric output never changes the live string field domain.
<Order.Slider name="amount" />;
// @ts-expect-error Bound validation ownership cannot be replaced.
Order.useForm({ resolver: async () => ({ values: { amount: 3 }, errors: {} }) });
// @ts-expect-error No flat adapter value escapes the compound family.
import { FormInput } from '@apvee/spfx-react-toolkit/controls/form';
// @ts-expect-error No plugin/advanced clean subpath is exported.
import {} from '@apvee/spfx-react-toolkit/controls/form/advanced';
void FormInput;
