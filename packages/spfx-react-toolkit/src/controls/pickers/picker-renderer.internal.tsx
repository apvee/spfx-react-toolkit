import * as React from 'react';
import { Avatar, Button, Field, Option, OptionGroup, Persona, Tag, TagPicker, TagPickerControl, TagPickerGroup, TagPickerInput, TagPickerList, TagPickerOption, TagPickerOptionGroup, Tooltip, useId } from '@fluentui/react-components';
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
import { PickerDropdown } from './picker-dropdown.internal';
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
  // React/Fluent DOM ref contracts use null while detached.
  const boundaryRef = React.useRef<HTMLDivElement>(null), popupRef = React.useRef<HTMLDivElement>(null), actionsRef = React.useRef<HTMLDivElement>(null);
  const blurVersion = React.useRef(0);
  React.useEffect(() => () => { blurVersion.current++; }, []);
  const contains = (target: EventTarget | null | undefined): boolean => {
    const owner = boundaryRef.current?.ownerDocument.defaultView;
    return !!(owner && target instanceof owner.Node && (boundaryRef.current?.contains(target) || popupRef.current?.contains(target)));
  };
  const onBlur: React.FocusEventHandler<HTMLDivElement> = event => {
    if (contains(event.relatedTarget)) {
      // Dismiss the overlapping native popup before pointerup on an owned action, without validating an exit.
      if (actionsRef.current?.contains(event.relatedTarget)) runtime.setOpen(false);
      return;
    }
    const version = ++blurVersion.current, owner = event.currentTarget.ownerDocument;
    const exit = (): void => { runtime.setOpen(false); void runtime.blur(); };
    const next = event.relatedTarget, ownerWindow = owner.defaultView;
    // Real outside controls must dismiss before pointerup; otherwise the floating popup can cover their click.
    if (ownerWindow && next instanceof ownerWindow.Element && next.closest('[aria-hidden="true"]') === null) { exit(); return; }
    // Native portal Tab navigation can cross aria-hidden intermediaries or omit its next target.
    ownerWindow?.setTimeout(() => {
      if (version === blurVersion.current && boundaryRef.current && !contains(owner.activeElement)) exit();
    }, 0);
  };
  const disabled = props.disabled || runtime.confirming;
  const typedSearch = props.searchable !== false && runtime.query.trim().length >= runtime.minimum;
  const available = (option: T): boolean => !runtime.selectedOptions.includes(option.key);
  const rows = typedSearch ? runtime.searchOptions.filter(available) : [];
  const pinnedSections = (runtime.pinnedSections ?? []).map(section => ({ ...section, options: section.options.filter(available) })).filter(section => section.options.length > 0);
  const sections = runtime.sections.map(section => ({ ...section, options: section.options.filter(available) })).filter(section => section.options.length > 0);
  const count = pinnedSections.reduce((sum, section) => sum + section.options.length, 0) + (typedSearch ? rows.length : sections.reduce((sum, section) => sum + section.options.length, 0));
  const avatar = (option: T, key: string): React.ReactElement => {
    const presentation = pickerPresentation(option), Icon = icons[presentation.icon];
    return <Avatar aria-hidden size={tag ? undefined : props.size === 'large' ? 24 : 20} name={presentation.user ? option.text : undefined} shape="square"
      color={tag && props.colorfulTagAvatars ? 'colorful' : presentation.user ? undefined : 'brand'} idForColor={tag && props.colorfulTagAvatars ? key : undefined}
      image={presentation.imageUrl === undefined ? undefined : { src: presentation.imageUrl }} icon={presentation.user ? undefined : <Icon />} />;
  };
  const optionNode = (option: T): React.ReactElement => {
    const presentation = pickerPresentation(option), Icon = icons[presentation.icon];
    const content = <Persona name={option.text} title={presentation.description} secondaryText={{ children: presentation.secondary, className: styles.text, title: presentation.secondary }}
      primaryText={{ className: styles.text, title: option.text }} className={styles.persona} size="extra-small" textAlignment="center"
      avatar={{ size: 24, shape: 'square', color: presentation.user ? undefined : 'brand', name: presentation.user ? option.text : undefined,
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
  const status = typedSearch ? searchStatus : belowMinimum ? props.emptyContent : runtime.loading ? props.loadingContent ?? 'Caricamento…' : runtime.error !== undefined ? props.errorContent ?? runtime.error.message : runtime.warning !== undefined ? runtime.warning :
    runtime.catalogHint !== undefined ? runtime.catalogHint : count === 0 ? props.emptyContent ?? (props.searchable === false ? 'Nessuna opzione disponibile.' : undefined) : undefined;
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
  const searchPlaceholder = props.placeholder ?? (props.searchable === false ? 'Cerca…' : prompt);
  const inputProps = { 'aria-label': name, 'aria-labelledby': props['aria-labelledby'], placeholder: searchPlaceholder };
  const selectionText = multi && runtime.selectedOptions.length > 1 ? `${runtime.selectedOptions.length} elementi selezionati` : runtime.selectedOptions.map(runtime.selectionText).join(', ');
  const selected = runtime.selectedOptions.length === 1 ? runtime.selectedItems.find(item => item.key === runtime.selectedOptions[0]) : undefined;
  const control = <div ref={boundaryRef} onBlur={onBlur} onFocus={() => { blurVersion.current++; }}><Field className={sx(styles.root, props.className)} label={props.id === undefined ? props.label : resolveShorthand(props.label, { defaultProps: { htmlFor: props.id } })} hint={props.hint} required={props.required}
    validationState={props.validationState ?? (runtime.validation ? 'error' : undefined)} validationMessage={props.validationMessage ?? runtime.validation}>
    {fieldProps => {
      const describedBy = [fieldProps['aria-describedby'], statusContent === undefined ? undefined : statusId].filter(Boolean).join(' ') || undefined;
      return <>{wrapControl(tag ? <TagPicker open={runtime.open} onOpenChange={(event, data) => { if (event.type !== 'blur') runtime.setOpen(data.open); }} disabled={props.disabled} size={props.size}
      selectedOptions={[...runtime.selectedOptions]} onOptionSelect={(event, data) => {
        if (disabled) return;
        if (runtime.selectedOptions.includes(data.value)) void runtime.remove(event, data.value);
        else void runtime.select(event, data.value);
      }}>
      <TagPickerControl aria-disabled={disabled || undefined} className={sx(styles.control, props.controlClassName)}>
        <TagPickerGroup className={styles.tagGroup} aria-label={name} aria-labelledby={props['aria-labelledby']}>
          {runtime.selectedOptions.map(key => { const option = runtime.selectedItems.find(candidate => candidate.key === key);
            return <Tag key={key} value={key} className={styles.tag} primaryText={{ className: styles.text, title: runtime.selectionText(key) }} dismissible disabled={disabled} media={option === undefined ? undefined : avatar(option, key)}>{runtime.selectionText(key)}</Tag>; })}
        </TagPickerGroup>
        <TagPickerInput {...fieldProps} {...inputProps} aria-describedby={describedBy} aria-disabled={disabled || undefined} autoComplete="off" id={props.id ?? fieldProps.id} value={runtime.query} readOnly={props.searchable === false || runtime.confirming}
          onChange={event => runtime.setQuery(event.currentTarget.value)} />
      </TagPickerControl>
      <TagPickerList ref={popupRef} className={styles.list}>{options}</TagPickerList>
    </TagPicker> : <PickerDropdown fieldProps={{ ...fieldProps, 'aria-label': name, 'aria-labelledby': props['aria-labelledby'], 'aria-describedby': describedBy, id: props.id ?? fieldProps.id }}
      runtime={runtime} multiselect={multi} disabled={props.disabled} searchable={props.searchable !== false} size={props.size}
      className={sx(styles.control, props.controlClassName)} selectionText={selectionText}
      selectionMedia={selected === undefined ? undefined : avatar(selected, selected.key)} searchPlaceholder={searchPlaceholder}
      name={name} contains={contains} popupRef={popupRef}>{options}</PickerDropdown>)}
    <div ref={actionsRef} className={styles.actions}>
      {showBlankOption ? <Button size="small" disabled={disabled} onClick={event => { void runtime.clear(event); }}>Nessuna selezione</Button> : undefined}
      {runtime.error !== undefined ? <Button size="small" disabled={disabled} onClick={runtime.retry}>Riprova</Button> : undefined}
      {runtime.canResolve ? <Button size="small" disabled={disabled || runtime.unresolvedOptions.length > 0} onClick={event => { void runtime.resolve(event); }}>Risolvi ID</Button> : undefined}
      {runtime.continuation !== undefined ? <Button size="small" disabled={disabled || runtime.loading} onClick={() => { void runtime.loadMore(); }}>Altri risultati</Button> : undefined}
    </div></>; }}
  </Field></div>;
  return props.tooltip === undefined ? control : <Tooltip {...props.tooltip}>{control}</Tooltip>;
}
