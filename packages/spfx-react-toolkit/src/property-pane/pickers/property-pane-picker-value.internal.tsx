import * as React from 'react';
import { useStableCallback } from '../../hooks/useStableCallback';
import { useSPFxPageContext } from '../../hooks/useSPFxPageContext';
import { getSPFxCorrelationInfo } from '../../helpers/spfx-page-context.helpers';
import type { PickerValue } from '../../controls/pickers/picker-values.types';
import type { PickerOption, PickerProps, PickerSelectionData, PickerSelectionEvent } from '../../controls/pickers/picker.types';
import { createPickerAlias, arePickerValuesCompatible } from '../../controls/pickers/picker-identity.internal';
import type { PickerIdentityContext } from '../../controls/pickers/picker-identity.internal';
import { copyPickerValue, validatePickerOption } from '../../controls/pickers/picker-values-validation.internal';
import { PickerSelectedValueContext, PickerAuthoritativeEpochContext } from '../../controls/pickers/picker-runtime-context.internal';

export type NativePickerChange = (targetProperty: string, value: readonly PickerValue[], isValidEntry: boolean) => void;
interface ValueBridgeProps<V extends PickerValue, T extends PickerOption<V>, C extends PickerProps<T>> {
  readonly Control: React.ComponentType<C>;
  readonly controlProps: C;
  readonly value: readonly V[];
  readonly valueRevision?: string | number;
  readonly targetProperty: string;
  readonly validate: (input: unknown) => readonly V[];
  readonly onChange?: NativePickerChange;
}
interface ValueAlias<V> { readonly key: string; readonly value: V; }
function valueAliases<V extends PickerValue>(values: readonly V[], previous: readonly ValueAlias<V>[], context: PickerIdentityContext): readonly ValueAlias<V>[] {
  const used = new Set<string>();
  return values.map(value => {
    const known = previous.find(entry => !used.has(entry.key) && arePickerValuesCompatible(entry.value, value, context));
    let key = known?.key ?? createPickerAlias(value);
    if (used.has(key)) {
      const base = key;
      let index = 1;
      while (used.has(key)) key = `${base}:duplicate:${index++}`;
    }
    used.add(key); return { key, value };
  });
}

/** Authoritative host references plus a pending draft for nonreactive panes; render/hydration never writes. */
export function PropertyPanePickerValueBridge<V extends PickerValue, T extends PickerOption<V>, C extends PickerProps<T> = PickerProps<T>>(props: ValueBridgeProps<V, T, C>): React.ReactElement {
  const page = useSPFxPageContext(), context = getSPFxCorrelationInfo(page);
  let values: readonly V[] = [], invalid: Error | undefined;
  try { values = props.validate(props.value); }
  catch (error) { invalid = error instanceof Error ? error : new Error('Invalid property picker references.'); }
  const signature = invalid === undefined ? JSON.stringify(values) : undefined;
  const [state, setState] = React.useState<{ readonly signature?: string; readonly revision?: string | number; readonly values: readonly V[]; readonly draft?: readonly V[]; readonly aliases: readonly ValueAlias<V>[] }>(() => ({
    signature, revision: props.valueRevision, values, aliases: invalid === undefined ? valueAliases(values, [], context) : [],
  }));
  const replace = invalid === undefined && (signature !== state.signature || props.valueRevision !== state.revision);
  const selection = replace ? values : state.draft ?? state.values;
  const aliases = valueAliases(selection, state.aliases, context);
  const refs = new Map(aliases.map(entry => [entry.key, entry.value]));
  React.useLayoutEffect(() => {
    if (replace) setState({ signature, revision: props.valueRevision, values, aliases });
  }, [signature, props.valueRevision, replace]);
  const [failure, setFailure] = React.useState<string>();
  React.useLayoutEffect(() => { if (replace) setFailure(undefined); }, [replace, signature, props.valueRevision]);
  const onSelect = useStableCallback((_event: PickerSelectionEvent, data: PickerSelectionData<T>): void => {
    if (invalid !== undefined || props.onChange === undefined || props.controlProps.disabled) return;
    try {
      if (!['selection', 'remove', 'clear', 'resolve'].includes(data.reason)) throw new Error('Value changes require an explicit picker action.');
      if (new Set(data.selectedOptions).size !== data.selectedOptions.length) throw new Error('Duplicate selection aliases.');
      const models = new Map(data.selectedItems.map(option => [option.key, option]));
      const next = props.validate(data.selectedOptions.map(key => {
        const retained = refs.get(key), model = models.get(key);
        // Ordinary actions retain original confirmed references, even when live models enrich them.
        if (retained !== undefined && data.reason !== 'resolve') return copyPickerValue(retained);
        if (model !== undefined) {
          if (retained !== undefined && !arePickerValuesCompatible(retained, model.value, context)) throw new Error('Resolution cannot replace a retained identity.');
          return copyPickerValue(model.value);
        }
        if (retained !== undefined) return copyPickerValue(retained);
        throw new Error('A new selection has no confirmed JSON reference.');
      }));
      setFailure(undefined);
      setState(previous => ({ ...previous, draft: next, aliases: valueAliases(next, aliases, context) }));
      props.onChange(props.targetProperty, next, true);
    } catch (error) { setFailure(error instanceof Error ? error.message : 'Invalid picker change.'); }
  });
  const withSelectedAlias = (option: T): T => {
    // Keep invalid models on the runtime's existing validation/reporting path.
    try { validatePickerOption(option); } catch { return option; }
    const alias = aliases.find(entry => arePickerValuesCompatible(entry.value, option.value, context));
    return alias === undefined ? option : { ...option, key: alias.key };
  };
  const initialOptions = props.controlProps.initialOptions?.map(withSelectedAlias);
  const availableOptions = props.controlProps.availableOptions?.map(withSelectedAlias);
  const missingCallback = props.onChange === undefined ? 'La callback nativa del Property Pane non è disponibile.' : undefined;
  const Control = props.Control;
  const controlProps: C = { ...props.controlProps, initialOptions, availableOptions, selectedOptions: aliases.map(entry => entry.key), onOptionSelect: onSelect,
    disabled: props.controlProps.disabled || invalid !== undefined || props.onChange === undefined,
    validationMessage: invalid?.message ?? failure ?? missingCallback ?? props.controlProps.validationMessage };
  return <PickerAuthoritativeEpochContext.Provider value={JSON.stringify([signature, props.valueRevision])}>
    <PickerSelectedValueContext.Provider value={refs}>
    <Control {...controlProps} />
    </PickerSelectedValueContext.Provider>
  </PickerAuthoritativeEpochContext.Provider>;
}
