import * as React from 'react';
import { createForm } from '../../packages/spfx-react-toolkit/src/controls/form';
import type { FormDatePickerStrings, FormDateTimeValue, FormTimePickerParseResult } from '../../packages/spfx-react-toolkit/src/controls/form';
const F = createForm<{ date: Date | null; time?: Date | null; onlyDate: Date; wrong: Date | null | string; appointment: FormDateTimeValue; rows: { appointment: FormDateTimeValue }[] }>();
const input = React.createRef<HTMLInputElement>();
const group = React.createRef<HTMLDivElement>();
const button = React.createRef<HTMLButtonElement>();
const strings: FormDatePickerStrings = { months: ['January'], shortMonths: ['Jan'], days: ['Sunday'], shortDays: ['Sun'], goToToday: 'Today' };
<F.DatePicker name="date" ref={input} strings={strings} onSelectDate={date => { const nullable: Date | null = date; void nullable; }} onBlur={event => { const target: HTMLInputElement = event.currentTarget; void target; }} parseDateFromString={() => null} />;
<F.TimePicker name="time" ref={input} parseTimeStringToDate={text => { const result: FormTimePickerParseResult = { date: text ? new Date(1970, 0, 1, 9) : null }; return result; }} onTimeChange={(event, data) => { const target: HTMLElement = event.currentTarget; const value: Date | null = data.selectedTime; void target; void value; }} />;
<F.DateTime name="appointment" ref={group} datePickerProps={{ allowTextInput: true }} timePickerProps={{ freeform: true, hourCycle: 'h23' }} />;
<F.Array name="rows">{({ items }) => items.map(({ Fields, key }) => <Fields.DateTime key={key} name="appointment" ref={group} />)}</F.Array>;
const methods = F.useForm({ defaultValues: { date: null, time: null, appointment: { date: null, time: null } } });
const actual: Date | null = methods.getValues('date'); void actual;
// @ts-expect-error clear cannot target nonnullable Date
<F.DatePicker name="onlyDate" />;
// @ts-expect-error native Date domain cannot read strings
<F.TimePicker name="wrong" />;
// @ts-expect-error controlled DatePicker defaults belong to RHF
<F.DatePicker name="date" value={new Date()} />;
// @ts-expect-error visible draft is owned by the adapter
<F.TimePicker name="time" value="12:00" />;
// @ts-expect-error independent time defaults belong to RHF
<F.TimePicker name="time" defaultSelectedTime={new Date()} />;
// @ts-expect-error no implicit string codec exists
<F.DatePicker name="date" codec={{}} />;
// @ts-expect-error DateTime public ref is the actual semantic group
<F.DateTime name="appointment" ref={button} />;
// @ts-expect-error DatePicker public ref is the actual native input
<F.DatePicker name="date" ref={group} />;
// @ts-expect-error unknown public time-validation error
<F.TimePicker name="time" parseTimeStringToDate={() => ({ date: null, error: 'bad' })} />;
// @ts-expect-error compat 0.4.21 does not support Input-only shadow appearances
<F.TimePicker name="time" appearance="filled-darker-shadow" />;
