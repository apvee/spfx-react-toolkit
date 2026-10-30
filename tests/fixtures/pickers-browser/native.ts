import type { IPropertyPaneCustomFieldProps, IPropertyPaneField } from '@microsoft/sp-property-pane';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker, PropertyPaneViewPicker, PropertyPaneContentTypePicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker } from '@apvee/spfx-react-toolkit/property-pane/pickers';
import type { SitePickerValue, ListPickerValue, FieldPickerValue, ViewPickerValue, ContentTypePickerValue, ListItemPickerValue, PrincipalPickerValue, PickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
import { sites, lists, fields, views, contentTypes, listContentTypes, items, principals, type Name } from './data';
import { host, blockReads, releaseReads, requests } from './transport';

type Mode = 'unresolved' | 'label' | 'model' | 'availableModel' | 'mixed' | 'empty';
interface Properties { sites: SitePickerValue[]; lists: ListPickerValue[]; fields: FieldPickerValue[]; views: ViewPickerValue[]; contentTypes: ContentTypePickerValue[]; items: ListItemPickerValue[]; principals: PrincipalPickerValue[]; }
// This is the same simulated SPFx owner used by the gallery; only its test property bag is narrowed.
const instance = host as unknown as BaseClientSideWebPart<Properties>;
function selection<V extends PickerValue>(first: V, second: V, mode: Mode): readonly V[] {
  if (mode === 'empty') return [];
  if (mode === 'mixed') return [{ ...first, text: 'Retained display snapshot' }, second];
  return [mode === 'label' ? { ...first, text: 'Retained display snapshot' } : first];
}
export function createNativeFixture(target: HTMLElement) {
  let field: IPropertyPaneField<IPropertyPaneCustomFieldProps> | undefined;
  let current: { name: Name; mode: Mode; changed: boolean; scope: 'site' | 'list'; refreshKey: number; presentation: 'combobox' | 'tagPicker' } | undefined;
  const values: Record<string, readonly PickerValue[]> = {};
  const changes: { target: string; value: readonly PickerValue[]; valid: boolean }[] = [];
  const validationCalls: string[] = [];
  const validateSelection = (): undefined => { validationCalls.push(current?.name ?? 'disposed'); return undefined; };
  function render(): void {
    if (current === undefined) throw new Error('Native fixture is not mounted');
    const { name, mode, changed, scope, refreshKey, presentation } = current;
    const index = changed ? 2 : 0;
    const common = { instance, key: 'native-browser-field', label: `Native ${name}`, hint: 'Native hydration hint', debounceMs: 0, refreshKey, multiselect: mode === 'mixed', presentation, validationTrigger: 'blur' as const, validateSelection };
    switch (name) {
      case 'SitePicker': {
        const value = selection(sites[index].value, sites[1].value, mode); values.current = value;
        field = PropertyPaneSitePicker('sites', { ...common, value, mode: 'hub', initialOptions: mode === 'model' ? [sites[index]] : undefined, availableOptions: mode === 'availableModel' ? [sites[index]] : undefined }); break;
      }
      case 'ListPicker': {
        const value = selection(lists[index].value, lists[1].value, mode); values.current = value;
        field = PropertyPaneListPicker('lists', { ...common, value, initialOptions: mode === 'model' ? [lists[index]] : undefined, availableOptions: mode === 'availableModel' ? [lists[index]] : undefined }); break;
      }
      case 'FieldPicker': {
        const value = selection(fields[index].value, fields[1].value, mode); values.current = value;
        field = PropertyPaneFieldPicker('fields', { ...common, value, list: 'Projects', initialOptions: mode === 'model' ? [fields[index]] : undefined, availableOptions: mode === 'availableModel' ? [fields[index]] : undefined }); break;
      }
      case 'ContentTypePicker': {
        const options = scope === 'site' ? contentTypes : listContentTypes;
        const value = selection(options[index].value, options[1].value, mode); values.current = value;
        const props = { ...common, value, initialOptions: mode === 'model' ? [options[index]] : undefined, availableOptions: mode === 'availableModel' ? [options[index]] : undefined };
        field = scope === 'site' ? PropertyPaneContentTypePicker('contentTypes', { ...props, scope: 'site' }) : PropertyPaneContentTypePicker('contentTypes', { ...props, scope: 'list', list: 'Projects' }); break;
      }
      case 'ViewPicker': {
        const value = selection(views[index].value, views[1].value, mode); values.current = value;
        field = PropertyPaneViewPicker('views', { ...common, value, list: 'Projects', initialOptions: mode === 'model' ? [views[index]] : undefined, availableOptions: mode === 'availableModel' ? [views[index]] : undefined }); break;
      }
      case 'ListItemPicker': {
        const value = selection(items[index].value, items[1].value, mode); values.current = value;
        field = PropertyPaneListItemPicker('items', { ...common, value, list: 'Projects', textField: 'Title', initialOptions: mode === 'model' ? [items[index]] : undefined, availableOptions: mode === 'availableModel' ? [items[index]] : undefined }); break;
      }
      case 'PrincipalPicker': {
        const value = selection(principals[index].value, principals[1].value, mode); values.current = value;
        field = PropertyPanePrincipalPicker('principals', { ...common, value, principalTypes: ['user'], initialOptions: mode === 'model' ? [principals[index]] : undefined, availableOptions: mode === 'availableModel' ? [principals[index]] : undefined }); break;
      }
    }
    field.properties.onRender(target, undefined, (property, value, valid) => changes.push({ target: property, value, valid }));
  }
  return {
    requests, values, changes, validationCalls, blockReads, releaseReads,
    mount(name: Name, mode: Mode = 'unresolved', presentation: 'combobox' | 'tagPicker' = 'tagPicker') {
      this.dispose(); current = { name, mode, changed: false, scope: 'site', refreshKey: 0, presentation }; render();
    },
    update(kind: 'equivalent' | 'selection' | 'refresh' | 'scope') {
      if (current === undefined) throw new Error('Native fixture is not mounted');
      if (kind === 'selection') current.changed = !current.changed;
      if (kind === 'scope') current.scope = current.scope === 'site' ? 'list' : 'site';
      if (kind === 'refresh') current.refreshKey++;
      render();
    },
    dispose() { field?.properties.onDispose?.(target); field = undefined; current = undefined; },
  };
}
