import * as React from 'react';
import { PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import type { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import { SPFxWebPartProvider } from '../../core/provider-webpart';
import type { PickerValue } from '../../controls/pickers/picker-values.types';
import type { PickerOption, PickerProps } from '../../controls/pickers/picker.types';
import type { PropertyPanePickerProps } from './property-pane-picker.types';
import { PropertyPanePickerHost } from './property-pane-picker-host.internal';
import { PropertyPanePickerValueBridge } from './property-pane-picker-value.internal';
import { mountOrUpdatePickerField, disposePickerField } from './property-pane-picker.internal';

/** The six native factories share lifecycle and value semantics, with domain-specific validated JSON arrays. */
export function createPropertyPanePicker<P extends object, V extends PickerValue, T extends PickerOption<V>, C extends PickerProps<T>>(targetProperty: string,
  props: Pick<PropertyPanePickerProps<P, C, V>, 'instance' | 'key' | 'value' | 'valueRevision'>, Control: React.ComponentType<C>, validate: (input: unknown) => readonly V[], controlProps: C): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  if (!targetProperty || /[.[\]]/.test(targetProperty) || ['__proto__', 'constructor', 'prototype'].includes(targetProperty)) throw new Error('Picker targetProperty must be a safe direct property name.');
  if (!props.key?.trim()) throw new Error('Property picker key must be nonempty and stable.');
  const token = {};
  return {
    type: PropertyPaneFieldType.Custom, targetProperty,
    properties: {
      key: props.key,
      onRender(domElement, _context, changeCallback) {
        const { instance, key, value, valueRevision } = props;
        mountOrUpdatePickerField(domElement, { owner: instance, key, token, render: runtime =>
          <SPFxWebPartProvider instance={instance}>
            <PropertyPanePickerHost document={domElement.ownerDocument} runtime={runtime}>
              <PropertyPanePickerValueBridge<V, T, C> Control={Control} controlProps={controlProps} value={value} valueRevision={valueRevision}
                validate={validate} targetProperty={targetProperty} onChange={changeCallback} />
            </PropertyPanePickerHost>
          </SPFxWebPartProvider>,
        });
      },
      onDispose(domElement) { disposePickerField(token, domElement); },
    },
  };
}
