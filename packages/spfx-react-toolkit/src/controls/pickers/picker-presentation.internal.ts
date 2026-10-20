import type { PickerOption } from './picker.types';

export type PickerIconKind = 'site' | 'list' | 'view' | 'text' | 'choice' | 'number' | 'money' | 'date' | 'boolean' | 'user' | 'link' | 'field' | 'item' | 'file' | 'pdf' | 'folder' | 'group' | 'email';
export interface PickerPresentation { readonly icon: PickerIconKind; readonly secondary?: string; readonly description?: string; readonly imageUrl?: string; readonly user: boolean; }
/** Literal secondary text; display snapshots never determine an identity. */
export function pickerPresentation(option: PickerOption): PickerPresentation {
  const metadata = option as unknown as Readonly<Record<string, unknown>>;
  const value = option.value;
  switch (value.kind) {
    case 'site': return { icon: 'site', secondary: new URL(value.url).pathname, user: false };
    case 'list': return { icon: 'list', secondary: metadata.baseTemplate === 101 ? 'Raccolta documenti' : 'Lista SharePoint', user: false };
    case 'contentType': return { icon: 'field', secondary: typeof metadata.group === 'string' ? metadata.group : undefined, description: typeof metadata.description === 'string' ? metadata.description : undefined, user: false };
    case 'view': return { icon: 'view', secondary: metadata.personalView ? 'Vista personale' : 'Vista pubblica', user: false };
    case 'field': {
      const type = typeof metadata.typeAsString === 'string' ? metadata.typeAsString : 'Campo';
      const fieldIcons: Readonly<Record<string, PickerIconKind>> = { Text: 'text', Note: 'text', Choice: 'choice', MultiChoice: 'choice', Number: 'number', Currency: 'money', DateTime: 'date', Boolean: 'boolean', User: 'user', UserMulti: 'user', Lookup: 'link', LookupMulti: 'link', URL: 'link' };
      return { icon: fieldIcons[type] ?? 'field', secondary: `${value.internalName} · ${type}`, user: false };
    }
    case 'item': return { icon: 'item', secondary: `Elemento · ID ${value.itemId}`, user: false };
    case 'file': return { icon: value.extension?.toLowerCase() === 'pdf' ? 'pdf' : 'file', secondary: value.extension === undefined ? 'File' : `File ${value.extension.toUpperCase()}`, user: false };
    case 'folder': return { icon: 'folder', secondary: 'Cartella', user: false };
    case 'user': return { icon: 'user', secondary: typeof metadata.email === 'string' ? metadata.email : typeof metadata.userPrincipalName === 'string' ? metadata.userPrincipalName : value.loginName,
      imageUrl: typeof metadata.imageUrl === 'string' ? metadata.imageUrl : undefined, user: true };
    case 'sharePointGroup': return { icon: 'group', secondary: 'Gruppo SharePoint · SharePoint', user: false };
    case 'microsoft365Group': return { icon: 'group', secondary: 'Microsoft 365 · Entra', user: false };
    case 'securityGroup': return { icon: 'group', secondary: 'Gruppo di sicurezza · Entra', user: false };
    case 'distributionList': return { icon: 'group', secondary: 'Lista di distribuzione · Entra', user: false };
    case 'unvalidatedEmail': return { icon: 'email', secondary: 'Email non verificata', user: false };
  }
}
