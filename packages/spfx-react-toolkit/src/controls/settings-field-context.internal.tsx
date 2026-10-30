import * as React from 'react';

/** Presentation of the single committed built-in field in one settings row. */
export interface SettingsFieldTarget {
  owner: symbol;
  id: string;
  nativeLabel: boolean;
  required: boolean;
  disabled: boolean;
}
/** Neutral presentation bridge: no values, validation state or RHF dependency. */
export interface SettingsFieldContextValue {
  labelId: string;
  descriptionId?: string;
  register: (target: SettingsFieldTarget) => () => void;
}
// Local context identity only; no external registration or renderer allocation.
export const SettingsFieldContext = React.createContext<SettingsFieldContextValue | undefined>(undefined);

/** Deduplicate additive ARIA id relationships without inspecting the DOM. */
export function mergeSettingsFieldIds(...values: (string | undefined)[]): string | undefined {
  const ids = values.flatMap(value => value?.split(/\s+/).filter(Boolean) ?? []);
  return ids.length ? [...new Set(ids)].join(' ') : undefined;
}
