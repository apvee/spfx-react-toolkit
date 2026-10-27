import type { PickersStrings, PickersStringsOverride } from './pickers-strings.types';
import { nativePickersResource } from './localization/pickers-native.internal';
import { formatResourceTemplate, formatResourcePlural } from '../../localization/resource-formatters.internal';

const resource = nativePickersResource;
const nativeStrings: PickersStrings = {
  labels: {
    selectOne: resource.labels.selectOne,
    selectMany: resource.labels.selectMany,
    search: resource.labels.search,
    clearSelection: resource.labels.clearSelection,
    clearSearch: resource.labels.clearSearch,
    options: resource.labels.options,
    namedOptions: (name: string) => formatResourceTemplate(resource.labels.namedOptions, { name }),
    searchOptions: resource.labels.searchOptions,
    namedSearch: (name: string) => formatResourceTemplate(resource.labels.namedSearch, { name }),
    selectedOne: resource.labels.selectedOne,
    selectedMany: resource.labels.selectedMany,
    noSelection: resource.labels.noSelection,
    retry: resource.labels.retry,
    resolveIds: resource.labels.resolveIds,
    moreResults: resource.labels.moreResults,
    selectedCount: (count: number) => formatResourcePlural(resource.labels.selectedCount, resource.__locale, count),
    minimumSearchLength: (count: number) => formatResourcePlural(resource.labels.minimumSearchLength, resource.__locale, count),
  },
  status: {
    searching: resource.status.searching,
    loading: resource.status.loading,
    partialSearch: resource.status.partialSearch,
    noOtherResults: resource.status.noOtherResults,
    noResults: resource.status.noResults,
    noOptions: resource.status.noOptions,
    moreItemsAvailable: resource.status.moreItemsAvailable,
    partialSuggestions: resource.status.partialSuggestions,
    unresolvedValue: resource.status.unresolvedValue,
    partialHubs: resource.status.partialHubs,
    partialSiteCollections: resource.status.partialSiteCollections,
  },
  sections: {
    consumer: resource.sections.consumer,
    currentSite: resource.sections.currentSite,
    hubSite: resource.sections.hubSite,
    catalog: resource.sections.catalog,
    currentUser: resource.sections.currentUser,
    people: resource.sections.people,
    memberOf: resource.sections.memberOf,
  },
  metadata: {
    documentLibrary: resource.metadata.documentLibrary,
    sharePointList: resource.metadata.sharePointList,
    personalView: resource.metadata.personalView,
    publicView: resource.metadata.publicView,
    field: resource.metadata.field,
    fieldDescription: (internalName: string, type: string) => formatResourceTemplate(resource.metadata.fieldDescription, { internalName, type }),
    itemDescription: (id: number) => formatResourceTemplate(resource.metadata.itemDescription, { id }),
    file: resource.metadata.file,
    fileWithExtension: (extension: string) => formatResourceTemplate(resource.metadata.fileWithExtension, { extension }),
    folder: resource.metadata.folder,
    sharePointGroup: resource.metadata.sharePointGroup,
    microsoft365Group: resource.metadata.microsoft365Group,
    securityGroup: resource.metadata.securityGroup,
    distributionList: resource.metadata.distributionList,
    unvalidatedEmail: resource.metadata.unvalidatedEmail,
  },
  validation: {
    selectionLimitReached: resource.validation.selectionLimitReached,
    selectionLimitExceeded: resource.validation.selectionLimitExceeded,
    unresolvedValues: resource.validation.unresolvedValues,
    outOfScopeValues: resource.validation.outOfScopeValues,
    resolveRequiredUserIds: resource.validation.resolveRequiredUserIds,
  },
  errors: {
    candidateOutOfScope: resource.errors.candidateOutOfScope,
    incompleteCandidateResolution: resource.errors.incompleteCandidateResolution,
    resolutionFailed: resource.errors.resolutionFailed,
    graphUnavailable: resource.errors.graphUnavailable,
  },
};

/** Known-field merge retains explicit empty values and ignores extra JavaScript keys. */
function mergeFields<T extends object>(defaults: T, overrides?: Partial<T>): T {
  const result = { ...defaults };
  for (const key of Object.keys(defaults) as Array<keyof T>) {
    const value = overrides?.[key];
    if (value !== undefined) result[key] = value;
  }
  return result;
}
export function resolvePickersStrings(overrides?: PickersStringsOverride): PickersStrings {
  return {
    labels: mergeFields(nativeStrings.labels, overrides?.labels),
    status: mergeFields(nativeStrings.status, overrides?.status),
    sections: mergeFields(nativeStrings.sections, overrides?.sections),
    metadata: mergeFields(nativeStrings.metadata, overrides?.metadata),
    validation: mergeFields(nativeStrings.validation, overrides?.validation),
    errors: mergeFields(nativeStrings.errors, overrides?.errors),
  };
}
