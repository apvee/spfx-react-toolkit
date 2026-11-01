import { privateValue } from './private';

export function PropertyPaneListPicker(name: string): { name: string } {
  return { name: name || privateValue };
}
