import { resolvePickersStrings } from './pickers-strings.internal';
import type { PickersStrings } from './pickers-strings.types';
import type { PickerOption } from './picker.types';
import type { PickerValue } from './picker-values.types';
import { canonicalPickerIdentity, pickerIdentityAliases } from './picker-identity.internal';
import type { PickerIdentityContext } from './picker-identity.internal';
import { validatePickerOption } from './picker-values-validation.internal';

export interface PickerSuggestionSection<T extends PickerOption> {
  readonly source: 'consumer' | 'currentUser' | 'people' | 'memberOf' | 'catalog' | 'currentSite' | 'hubSite';
  readonly label: string;
  readonly options: readonly T[];
}
export interface PickerSuggestionInput<T extends PickerOption> {
  readonly strings?: PickersStrings;
  readonly consumer?: readonly T[];
  readonly catalog?: readonly T[];
  readonly catalogLimit?: number;
  readonly completeCatalog?: boolean;
  readonly currentSite?: readonly T[];
  readonly hubSite?: readonly T[];
  readonly currentUser?: readonly T[];
  readonly people?: readonly T[];
  readonly groups?: readonly T[];
  readonly selected?: readonly T[];
  readonly selectedValues?: readonly PickerValue[];
  readonly context?: PickerIdentityContext;
  readonly limit?: number;
  readonly showDefaultSuggestions?: boolean;
  readonly accepts?: (option: T) => boolean;
  readonly filterOptions?: (options: readonly T[]) => readonly T[];
}

/** Pure composition; failures in one external source do not erase consumer candidates. */
export function composePickerSuggestions<T extends PickerOption>(input: PickerSuggestionInput<T>): readonly PickerSuggestionSection<T>[] {
  const strings = input.strings ?? resolvePickersStrings();
  if (input.showDefaultSuggestions === false) return [];
  const limit = input.completeCatalog ? Number.MAX_SAFE_INTEGER : input.catalogLimit ?? input.limit ?? 5;
  if (!input.completeCatalog && (!Number.isSafeInteger(limit) || limit < 1 || limit > (input.catalogLimit === undefined ? 5 : 200))) throw new RangeError('Invalid initial suggestion/catalog limit.');
  const excluded = new Set([...(input.selected ?? []).map(option => option.value), ...(input.selectedValues ?? [])].flatMap(value => pickerIdentityAliases(value, input.context)));
  const sourceRows: readonly [PickerSuggestionSection<T>['source'], string, readonly T[]][] = [
    ['consumer', strings.sections.consumer, input.consumer ?? []],
    ['currentSite', strings.sections.currentSite, input.currentSite ?? []],
    ['hubSite', strings.sections.hubSite, input.hubSite ?? []],
    ['catalog', strings.sections.catalog, input.catalog ?? []],
    ['currentUser', strings.sections.currentUser, input.currentUser ?? []],
    ['people', strings.sections.people, input.people ?? []],
    ['memberOf', strings.sections.memberOf, input.groups ?? []],
  ];
  const sections: PickerSuggestionSection<T>[] = [];
  let count = 0;
  for (const [source, label, candidates] of sourceRows) {
    const valid = candidates.filter(option => {
      try {
        validatePickerOption(option);
        return (input.accepts?.(option) ?? true) && !pickerIdentityAliases(option.value, input.context).some(alias => excluded.has(alias));
      } catch { return false; }
    });
    const permitted = new Map(valid.map(option => [option.key, canonicalPickerIdentity(option.value, input.context)]));
    const transformed = input.filterOptions?.(valid) ?? valid;
    const options: T[] = [];
    for (const option of transformed) {
      try {
        validatePickerOption(option);
        if (permitted.get(option.key) !== canonicalPickerIdentity(option.value, input.context) || !(input.accepts?.(option) ?? true)) continue;
        const aliases = pickerIdentityAliases(option.value, input.context);
        if (aliases.some(alias => excluded.has(alias))) continue;
        if (count >= limit) break;
        aliases.forEach(alias => excluded.add(alias));
        options.push(option); count++;
      } catch { /* Invalid transformations cannot create an unverified candidate. */ }
    }
    if (options.length > 0) sections.push({ source, label, options });
    if (count >= limit) break;
  }
  return sections;
}
