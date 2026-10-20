import * as React from 'react';
import { Avatar, Button, Combobox, Field, Option, OptionGroup, Persona, Tag, TagPicker, TagPickerControl, TagPickerGroup, TagPickerInput, TagPickerList, TagPickerOption, TagPickerOptionGroup, Tooltip, useId } from '@fluentui/react-components';
import { AppsListRegular, CalendarRegular, CheckboxCheckedRegular, DocumentPdfRegular, DocumentRegular, EyeRegular, FolderRegular, GlobeRegular, LinkRegular, MailRegular, MoneyRegular, NumberSymbolRegular, OptionsRegular, PeopleTeamRegular, PersonRegular, TextAlignLeftRegular, TextBulletListSquareRegular } from '@fluentui/react-icons';
import { resolveShorthand } from '@fluentui/react-utilities';
import { useSx } from '../../hooks/useSx';
import { renderContentSlot } from '../content-slot.internal';
import type { PickerOption, PickerProps } from './picker.types';
import type { PickerValue } from './picker-values.types';
import type { usePickerRuntime } from './usePickerRuntime.internal';
import { pickerPresentation } from './picker-presentation.internal';
import type { PickerIconKind } from './picker-presentation.internal';
import { PickerControlBoundaryContext } from './picker-runtime-context.internal';
import { usePickerStyles } from './picker-styles.internal';

const icons: Readonly<Record<PickerIconKind, React.ComponentType>> = {
  site: GlobeRegular, list: AppsListRegular, view: EyeRegular, text: TextAlignLeftRegular, choice: TextBulletListSquareRegular,
  number: NumberSymbolRegular, money: MoneyRegular, date: CalendarRegular, boolean: CheckboxCheckedRegular, user: PersonRegular,
  link: LinkRegular, field: OptionsRegular, item: AppsListRegular, file: DocumentRegular, pdf: DocumentPdfRegular, folder: FolderRegular, group: PeopleTeamRegular, email: MailRegular,
};
export interface PickerRendererProps<V extends PickerValue, T extends PickerOption<V>> {
  readonly props: PickerProps<T>;
  readonly runtime: ReturnType<typeof usePickerRuntime<V, T>>;
  readonly defaultPresentation: 'combobox' | 'tagPicker';
  readonly defaultMultiselect: boolean;
  readonly showBlankOption?: boolean;
}

/** Native Fluent renderer; semantic labels/status/actions remain control owned. */
export function PickerRenderer<V extends PickerValue, T extends PickerOption<V>>({ props, runtime, defaultPresentation, defaultMultiselect, showBlankOption }: PickerRendererProps<V, T>): React.ReactElement {
  const styles = usePickerStyles(), sx = useSx(), statusId = useId('picker-status');
  const ControlBoundary = React.useContext(PickerControlBoundaryContext);
  const wrapControl = (children: React.ReactElement): React.ReactElement => ControlBoundary === undefined ? children :
    <ControlBoundary pending={runtime.initialHydrationPending}>{children}</ControlBoundary>;
  const tag = (props.presentation ?? defaultPresentation) === 'tagPicker';
  const multi = props.multiselect ?? defaultMultiselect;
  const name = props['aria-label'] ?? (typeof props.label === 'string' ? props.label : undefined);
  if (!name?.trim() && !props['aria-labelledby']) throw new TypeError('Picker requires a plain label, aria-label or aria-labelledby.');
  const disabled = props.disabled || runtime.confirming;
  const typedSearch = props.searchable !== false && runtime.query.trim().length >= runtime.minimum;
  const available = (option: T): boolean => tag || !runtime.selectedOptions.includes(option.key);
  const rows = typedSearch ? runtime.searchOptions.filter(available) : [];
  const pinnedSections = (runtime.pinnedSections ?? []).map(section => ({ ...section, options: section.options.filter(available) })).filter(section => section.options.length > 0);
  const sections = runtime.sections.map(section => ({ ...section, options: section.options.filter(available) })).filter(section => section.options.length > 0);
  const count = pinnedSections.reduce((sum, section) => sum + section.options.length, 0) + (typedSearch ? rows.length : sections.reduce((sum, section) => sum + section.options.length, 0));
  const avatar = (option: T): React.ReactElement => {
    const presentation = pickerPresentation(option), Icon = icons[presentation.icon];
    return <Avatar aria-hidden name={presentation.user ? option.text : undefined} shape={presentation.user ? 'circular' : 'square'}
      image={presentation.imageUrl === undefined ? undefined : { src: presentation.imageUrl }} icon={presentation.user ? undefined : <Icon />} />;
  };
  const optionNode = (option: T): React.ReactElement => {
    const presentation = pickerPresentation(option), Icon = icons[presentation.icon];
    const content = <Persona name={option.text} title={presentation.description} secondaryText={{ children: presentation.secondary, className: styles.text, title: presentation.secondary }}
      primaryText={{ className: styles.text, title: option.text }} className={styles.persona} size="extra-small" textAlignment="center"
      avatar={{ shape: presentation.user ? 'circular' : 'square', name: presentation.user ? option.text : undefined,
        image: presentation.imageUrl === undefined ? undefined : { src: presentation.imageUrl }, icon: presentation.user ? undefined : <Icon /> }} />;
    return tag ? <TagPickerOption key={option.key} value={option.key} text={option.text}>{content}</TagPickerOption> :
      <Option key={option.key} value={option.key} text={option.text}
        // Fluent requires null to omit the check column for unselected single options; selected rows retain the native check.
        checkIcon={multi || runtime.selectedOptions.includes(option.key) ? undefined : null}>{content}</Option>;
  };
  const prompt = `Digita almeno ${runtime.minimum} caratteri per cercare`;
  const feedback = runtime.searchFeedback;
  const searchStatus = feedback.phase === 'pending' ? props.loadingContent ?? 'Ricerca in corso…' :
    feedback.phase === 'error' ? props.errorContent ?? feedback.error?.message :
    feedback.phase === 'partial' ? feedback.warning ?? 'Ricerca parziale: affina la ricerca.' :
    feedback.phase === 'complete' && rows.length === 0 ? props.emptyContent ?? (feedback.matchesVisibleChoices ? 'Nessun altro risultato per questa ricerca' : 'Nessun risultato per questa ricerca') : undefined;
  const belowMinimum = props.searchable !== false && runtime.query.trim().length > 0 && !typedSearch;
  const status = typedSearch ? searchStatus : belowMinimum ? props.emptyContent ?? prompt : runtime.loading ? props.loadingContent ?? 'Caricamento…' : runtime.error !== undefined ? props.errorContent ?? runtime.error.message : runtime.warning !== undefined ? runtime.warning :
    runtime.catalogHint !== undefined ? runtime.catalogHint : count === 0 ? props.emptyContent ?? (props.searchable === false ? 'Nessuna opzione disponibile.' : prompt) : undefined;
  // Contextual metadata failures do not replace feedback for the active typed query.
  const queryFeedback = typedSearch || belowMinimum;
  const additionalError = queryFeedback ? runtime.contextualError ?? (runtime.error !== feedback.error ? runtime.error : undefined) : undefined;
  const additionalStatus = additionalError !== undefined ? props.errorContent ?? additionalError.message : queryFeedback ? runtime.contextualWarning : undefined;
  const statusContent = status === undefined && additionalStatus === undefined ? undefined : <div id={statusId} role="status" className={styles.status}>
    {status === undefined ? undefined : renderContentSlot(status, 'div', sx, { shorthand: 'preserve' })}
    {additionalStatus === undefined ? undefined : renderContentSlot(additionalStatus, 'div', sx, { shorthand: 'preserve' })}
  </div>;
  const options = <>
    {props.suggestionsHeader === undefined ? undefined : <div role="presentation" className={styles.status}>{renderContentSlot(props.suggestionsHeader, 'span', sx, { shorthand: 'preserve' })}</div>}
    {!tag && runtime.selectedOptions.length > 0 ? <OptionGroup label={multi ? 'Selezionati' : 'Selezionato'}>{runtime.selectedOptions.map(key => {
      const option = runtime.selectedItems.find(item => item.key === key);
      return option === undefined ? <Option key={key} value={key} text={runtime.selectionText(key)}>{runtime.selectionText(key)}</Option> : optionNode(option);
    })}</OptionGroup> : undefined}
    {pinnedSections.map(section => tag ?
      <TagPickerOptionGroup key={section.source} label={section.label}>{section.options.map(optionNode)}</TagPickerOptionGroup> :
      <OptionGroup key={section.source} label={section.label}>{section.options.map(optionNode)}</OptionGroup>)}
    {typedSearch ? rows.map(optionNode) : sections.map(section => tag ?
      <TagPickerOptionGroup key={section.source} label={section.label}>{section.options.map(optionNode)}</TagPickerOptionGroup> :
      <OptionGroup key={section.source} label={section.label}>{section.options.map(optionNode)}</OptionGroup>)}
    {statusContent}
  </>;
  const inputProps = { 'aria-label': name, 'aria-labelledby': props['aria-labelledby'], 'aria-describedby': statusContent === undefined ? undefined : statusId,
    placeholder: props.placeholder ?? 'Cerca…', onBlur: () => { void runtime.blur(); } };
  const selectionText = runtime.selectedOptions.map(runtime.selectionText).join(', ');
  const control = <Field className={sx(styles.root, props.className)} label={props.id === undefined ? props.label : resolveShorthand(props.label, { defaultProps: { htmlFor: props.id } })} hint={props.hint} required={props.required}
    validationState={props.validationState ?? (runtime.validation ? 'error' : undefined)} validationMessage={props.validationMessage ?? runtime.validation}>
    {fieldProps => <>{wrapControl(tag ? <TagPicker open={runtime.open} onOpenChange={(_, data) => runtime.setOpen(data.open)} disabled={disabled} size={props.size}
      selectedOptions={[...runtime.selectedOptions]} onOptionSelect={(event, data) => {
        if (runtime.selectedOptions.includes(data.value)) void runtime.remove(event, data.value);
        else void runtime.select(event, data.value);
      }}>
      <TagPickerControl className={sx(styles.control, props.controlClassName)}>
        <TagPickerGroup className={styles.tagGroup} aria-label={name} aria-labelledby={props['aria-labelledby']}>
          {runtime.selectedOptions.map(key => { const option = runtime.selectedItems.find(candidate => candidate.key === key);
            return <Tag key={key} value={key} className={styles.tag} primaryText={{ className: styles.text, title: runtime.selectionText(key) }} dismissible disabled={disabled} media={option === undefined ? undefined : avatar(option)}>{runtime.selectionText(key)}</Tag>; })}
        </TagPickerGroup>
        <TagPickerInput {...fieldProps} {...inputProps} autoComplete="off" id={props.id ?? fieldProps.id} value={runtime.query} readOnly={props.searchable === false}
          onChange={event => runtime.setQuery(event.currentTarget.value)} />
      </TagPickerControl>
      <TagPickerList className={styles.list}>{options}</TagPickerList>
    </TagPicker> : <Combobox {...fieldProps} {...inputProps} id={props.id ?? fieldProps.id} className={sx(styles.control, props.controlClassName)} disabled={disabled} size={props.size}
      // Fluent autoSize writes an inline max-height that overrides the shared bounded popup class.
      positioning={{ autoSize: false }} listbox={{ className: styles.list }} clearable={!multi} clearIcon={{ 'aria-label': 'Cancella selezione' }} multiselect={multi} selectedOptions={[...runtime.selectedOptions]} open={runtime.open} onOpenChange={(_, data) => runtime.setOpen(data.open)}
      value={runtime.open ? runtime.query : selectionText} input={{ autoComplete: 'off', readOnly: props.searchable === false }}
      onChange={event => runtime.setQuery(event.currentTarget.value)} onOptionSelect={(event, data) => {
        if (data.optionValue !== undefined) {
          if (runtime.selectedOptions.includes(data.optionValue)) {
            if (multi) void runtime.remove(event, data.optionValue);
          } else void runtime.select(event, data.optionValue);
        }
        // Fluent clears single selection on unmatched input too; typing changes only the query.
        else if (data.selectedOptions.length === 0 && event.type !== 'change' && event.type !== 'input') void runtime.clear(event);
      }}>
      {options}
    </Combobox>)}
    <div className={styles.actions}>
      {showBlankOption ? <Button size="small" disabled={disabled} onClick={event => { void runtime.clear(event); }}>Nessuna selezione</Button> : undefined}
      {runtime.error !== undefined ? <Button size="small" disabled={disabled} onClick={runtime.retry}>Riprova</Button> : undefined}
      {runtime.canResolve ? <Button size="small" disabled={disabled || runtime.unresolvedOptions.length > 0} onClick={event => { void runtime.resolve(event); }}>Risolvi ID</Button> : undefined}
      {runtime.continuation !== undefined ? <Button size="small" disabled={disabled || runtime.loading} onClick={() => { void runtime.loadMore(); }}>Altri risultati</Button> : undefined}
    </div></>}
  </Field>;
  return props.tooltip === undefined ? control : <Tooltip {...props.tooltip}>{control}</Tooltip>;
}
