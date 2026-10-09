import * as React from 'react';
import { Input, type InputProps } from '@fluentui/react-components';
import { DatePicker, type DatePickerProps } from '@fluentui/react-datepicker-compat';
import { TimePicker, type TimePickerProps } from '@fluentui/react-timepicker-compat';
import { AddRegular } from '@fluentui/react-icons';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import type { StandardSchemaV1 } from '@standard-schema/spec';
import { useForm, useController, useFieldArray, type Resolver } from 'react-hook-form';

type InputValues = { amount: string; rows: { label: string }[] };
type OutputValues = { amount: number; rows: { label: string }[] };
const schema: StandardSchemaV1<InputValues, OutputValues> = {
  '~standard': {
    version: 1,
    vendor: 'dependency-probe',
    validate: async value => {
      if (typeof value !== 'object' || value === null || !('amount' in value) ||
          typeof value.amount !== 'string' || !('rows' in value) || !Array.isArray(value.rows)) {
        return { issues: [{ message: 'Invalid input' }] };
      }
      return { value: { amount: Number(value.amount), rows: [] } };
    }
  }
};
const resolver: Resolver<InputValues, undefined, OutputValues> = standardSchemaResolver(schema);

export function DependencyTypes(): React.ReactElement {
  const form = useForm<InputValues, undefined, OutputValues>({ resolver });
  const controller = useController({ control: form.control, name: 'amount' });
  const rows = useFieldArray({ control: form.control, name: 'rows' });
  const input: string = form.getValues('amount');
  rows.append({ label: input });
  // @ts-expect-error Input paths remain constrained to the input domain.
  form.setValue('amount', 42);
  // @ts-expect-error Array rows retain their input shape.
  rows.append({ label: 42 });
  const submit = form.handleSubmit(output => {
    const amount: number = output.amount;
    // @ts-expect-error Validation transforms only the submission output.
    const text: string = output.amount;
    void amount; void text;
  });
  const inputRef = React.createRef<HTMLInputElement>();
  const onChange: NonNullable<InputProps['onChange']> = (event, data) => {
    const target: HTMLInputElement = event.currentTarget;
    const value: string = data.value;
    controller.field.onChange(value);
    void target;
  };
  const onSelectDate: DatePickerProps['onSelectDate'] = date => {
    const selected: Date | null | undefined = date;
    void selected;
  };
  const onTimeChange: TimePickerProps['onTimeChange'] = (event, data) => {
    const selected: Date | null = data.selectedTime;
    const text: string | undefined = data.selectedTimeText;
    void event; void selected; void text;
  };
  return <form onSubmit={submit}>
    <Input ref={inputRef} value={controller.field.value} onChange={onChange} />
    <DatePicker ref={inputRef} onSelectDate={onSelectDate} />
    <TimePicker ref={inputRef} onTimeChange={onTimeChange} />
    <AddRegular />
  </form>;
}

/** Identity probe for the shared React Hook Form webpack external. */
export const dependencyIdentity = { useForm, useController, useFieldArray };
