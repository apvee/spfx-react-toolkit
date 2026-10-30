/** Translator-authored cardinal forms; the generator requires the exact categories for each culture. */
interface PickersPluralResource {
  readonly zero?: string;
  readonly one?: string;
  readonly two?: string;
  readonly few?: string;
  readonly many?: string;
  readonly other: string;
}

/** Private, data-only source contract. Named templates become callbacks in the family adapter. */
export interface PickersResource {
  readonly labels: {
    readonly selectOne: string;
    readonly selectMany: string;
    readonly search: string;
    readonly clearSelection: string;
    readonly clearSearch: string;
    readonly options: string;
    /** Template token: name. */
    readonly namedOptions: string;
    readonly searchOptions: string;
    /** Template token: name. */
    readonly namedSearch: string;
    readonly selectedOne: string;
    readonly selectedMany: string;
    readonly noSelection: string;
    readonly retry: string;
    readonly resolveIds: string;
    readonly moreResults: string;
    /** Cardinal templates, each containing count exactly once. */
    readonly selectedCount: PickersPluralResource;
    /** Cardinal templates, each containing count exactly once. */
    readonly minimumSearchLength: PickersPluralResource;
  };
  readonly status: {
    readonly searching: string;
    readonly loading: string;
    readonly partialSearch: string;
    readonly noOtherResults: string;
    readonly noResults: string;
    readonly noOptions: string;
    readonly moreItemsAvailable: string;
    readonly partialSuggestions: string;
    readonly unresolvedValue: string;
    readonly partialHubs: string;
    readonly partialSiteCollections: string;
  };
  readonly sections: {
    readonly consumer: string;
    readonly currentSite: string;
    readonly hubSite: string;
    readonly catalog: string;
    readonly currentUser: string;
    readonly people: string;
    readonly memberOf: string;
  };
  readonly metadata: {
    readonly documentLibrary: string;
    readonly sharePointList: string;
    readonly personalView: string;
    readonly publicView: string;
    readonly field: string;
    /** Template tokens: internalName, type. Both values remain source data. */
    readonly fieldDescription: string;
    /** Template token: id. */
    readonly itemDescription: string;
    readonly file: string;
    /** Template token: extension. */
    readonly fileWithExtension: string;
    readonly folder: string;
    readonly sharePointGroup: string;
    readonly microsoft365Group: string;
    readonly securityGroup: string;
    readonly distributionList: string;
    readonly unvalidatedEmail: string;
  };
  readonly validation: {
    readonly selectionLimitReached: string;
    readonly selectionLimitExceeded: string;
    readonly unresolvedValues: string;
    readonly outOfScopeValues: string;
    readonly resolveRequiredUserIds: string;
  };
  readonly errors: {
    readonly candidateOutOfScope: string;
    readonly incompleteCandidateResolution: string;
    readonly resolutionFailed: string;
    readonly graphUnavailable: string;
  };
}
