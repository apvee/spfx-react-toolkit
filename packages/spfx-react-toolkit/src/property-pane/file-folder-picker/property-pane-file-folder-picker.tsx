import * as React from 'react';
import { PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import type { IPropertyPaneCustomFieldProps, IPropertyPaneField } from '@microsoft/sp-property-pane';
import { SPFxWebPartProvider } from '../../core/provider-webpart';
import { PropertyPanePickerHost } from '../pickers/property-pane-picker-host.internal';
import { disposePickerField, mountOrUpdatePickerField } from '../pickers/property-pane-picker.internal';
import { PropertyPaneFileFolderPickerValueBridge } from './property-pane-file-folder-picker-value.internal';
import type { FileFolderPickerPropertyKey, PropertyPaneFileFolderPickerProps } from './property-pane-file-folder-picker.types';

/** Native SPFx property-pane field for existing files/folders through the hosted Microsoft File Picker v8.
 * Requires an initialized web part and existing delegated permissions. React FileFolderPicker options/defaults apply,
 * while SPFx owns controlled JSON arrays and apply/persistence. Every accepted action emits targetProperty, fresh
 * validated FileFolderPickerValue[] and true; hydration, display-model changes and authoritative resets never emit.
 * Equivalent renders retain provider/root/cache and nonreactive draft. valueRevision resets the draft and cancels
 * pending confirmation. Without a native change callback the field displays saved values with mutation disabled.
 * Initial unlabeled selections use the existing input-only AutoSkeleton boundary; known labels remain visible.
 * Latest-factory disposal closes owned picker resources and releases its metadata lease; stale disposal is ignored.
 * @typeParam TProperties Web-part property bag containing an exact file/folder reference array.
 * @param targetProperty Safe direct array property key; nested paths and prototype names are rejected at runtime too.
 * @param props Initialized instance, stable field key, required JSON array, optional revision and React control options.
 * @returns Native custom field using the owning document's renderer, native pane token theme and SPFxWebPartProvider.
 * @throws Invalid targetProperty or blank field key.
 * @example
 * PropertyPaneFileFolderPicker('documents', {
 *   instance: this, key: 'documents-picker', value: this.properties.documents ?? [],
 *   label: 'Documents', multiselect: true, scope: { kind: 'siteCollection' },
 * });
 */
export function PropertyPaneFileFolderPicker<TProperties extends object>(
  targetProperty: FileFolderPickerPropertyKey<TProperties>, props: PropertyPaneFileFolderPickerProps<TProperties>,
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  if (typeof targetProperty !== 'string' || !targetProperty || /[.[\]]/.test(targetProperty) || ['__proto__', 'constructor', 'prototype'].includes(targetProperty)) {
    throw new Error('FileFolderPicker targetProperty must be a safe direct property name.');
  }
  if (typeof props.key !== 'string' || !props.key.trim()) throw new Error('FileFolderPicker key must be nonempty and stable.');
  const { instance, key, value, valueRevision, ...controlProps } = props;
  const token = {};
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key,
      onRender(domElement, _context, changeCallback) {
        mountOrUpdatePickerField(domElement, { owner: instance, key, token, render: runtime =>
          <SPFxWebPartProvider instance={instance}>
            <PropertyPanePickerHost container={domElement} runtime={runtime}>
              <PropertyPaneFileFolderPickerValueBridge controlProps={controlProps} value={value} valueRevision={valueRevision}
                targetProperty={targetProperty} onChange={changeCallback} />
            </PropertyPanePickerHost>
          </SPFxWebPartProvider>,
        });
      },
      onDispose(domElement) { disposePickerField(token, domElement); },
    },
  };
}
