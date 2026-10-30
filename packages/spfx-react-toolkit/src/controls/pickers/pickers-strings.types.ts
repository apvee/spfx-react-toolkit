/** Complete toolkit-owned Pickers UI dictionary. Native SPFx resources supply defaults.
 * Caller labels, content slots, values and backend errors keep their own text.
 * Callback results are plain text; source arguments are passed literally without template reprocessing.
 * @example
 * ```ts
 * function customize(strings: PickersStrings): PickersStrings {
 *   return { ...strings, status: { ...strings.status, loading: 'Please wait' } };
 * }
 * ```
 */
export interface PickersStrings {
  /** Labels text. Every member is required in a complete dictionary. */
  readonly labels: {
    /** English native default: Select an item */
    readonly selectOne: string;
    /** English native default: Select items */
    readonly selectMany: string;
    /** English native default: Search… */
    readonly search: string;
    /** English native default: Clear selection */
    readonly clearSelection: string;
    /** English native default: Clear search */
    readonly clearSearch: string;
    /** English native default: Options */
    readonly options: string;
    /** Caller control name, inserted literally. Returns the visible/accessibility text.
     * @param name - Caller control name.
     * @returns Plain UI text.
     */
    readonly namedOptions: (name: string) => string;
    /** English native default: Search options */
    readonly searchOptions: string;
    /** Caller control name, inserted literally. Returns the visible/accessibility text.
     * @param name - Caller control name.
     * @returns Plain UI text.
     */
    readonly namedSearch: (name: string) => string;
    /** English native default: Selected item */
    readonly selectedOne: string;
    /** English native default: Selected items */
    readonly selectedMany: string;
    /** English native default: No selection */
    readonly noSelection: string;
    /** English native default: Retry */
    readonly retry: string;
    /** English native default: Resolve IDs */
    readonly resolveIds: string;
    /** English native default: More results */
    readonly moreResults: string;
    /** Number of selected items; native resource culture selects plural grammar. Returns the visible/accessibility text.
     * @param count - Number of selected items.
     * @returns Plain UI text.
     */
    readonly selectedCount: (count: number) => string;
    /** Minimum search character count; native resource culture selects plural grammar. Returns the visible/accessibility text.
     * @param count - Minimum searchable character count.
     * @returns Plain UI text.
     */
    readonly minimumSearchLength: (count: number) => string;
  };
  /** Status text. Every member is required in a complete dictionary. */
  readonly status: {
    /** English native default: Searching… */
    readonly searching: string;
    /** English native default: Loading… */
    readonly loading: string;
    /** English native default: Partial search: refine your search. */
    readonly partialSearch: string;
    /** English native default: No other results for this search */
    readonly noOtherResults: string;
    /** English native default: No results for this search */
    readonly noResults: string;
    /** English native default: No options available. */
    readonly noOptions: string;
    /** English native default: More items are available. */
    readonly moreItemsAvailable: string;
    /** English native default: Partial suggestions: some sources are unavailable. */
    readonly partialSuggestions: string;
    /** English native default: Unresolved value */
    readonly unresolvedValue: string;
    /** English native default: Partial hub list: narrow the scope or use site mode. */
    readonly partialHubs: string;
    /** English native default: Partial site collection search (maximum 1000 results): narrow your criteria. */
    readonly partialSiteCollections: string;
  };
  /** Sections text. Every member is required in a complete dictionary. */
  readonly sections: {
    /** English native default: Provided options */
    readonly consumer: string;
    /** English native default: Current site */
    readonly currentSite: string;
    /** English native default: Associated hub */
    readonly hubSite: string;
    /** English native default: Available options */
    readonly catalog: string;
    /** English native default: You */
    readonly currentUser: string;
    /** English native default: Suggested people */
    readonly people: string;
    /** English native default: Your groups */
    readonly memberOf: string;
  };
  /** Metadata text. Every member is required in a complete dictionary. */
  readonly metadata: {
    /** English native default: Document library */
    readonly documentLibrary: string;
    /** English native default: SharePoint list */
    readonly sharePointList: string;
    /** English native default: Personal view */
    readonly personalView: string;
    /** English native default: Public view */
    readonly publicView: string;
    /** English native default: Field */
    readonly field: string;
    /** Source field internal name and type, inserted literally. Returns the visible/accessibility text.
     * @param internalName - Source field internal name.
     * @param type - Source field type name; never translated.
     * @returns Plain UI text.
     */
    readonly fieldDescription: (internalName: string, type: string) => string;
    /** Source item ID, inserted literally. Returns the visible/accessibility text.
     * @param id - Source item ID.
     * @returns Plain UI text.
     */
    readonly itemDescription: (id: number) => string;
    /** English native default: File */
    readonly file: string;
    /** Source uppercase file extension, inserted literally. Returns the visible/accessibility text.
     * @param extension - Uppercase source extension.
     * @returns Plain UI text.
     */
    readonly fileWithExtension: (extension: string) => string;
    /** English native default: Folder */
    readonly folder: string;
    /** English native default: SharePoint group · SharePoint */
    readonly sharePointGroup: string;
    /** English native default: Microsoft 365 · Entra */
    readonly 'microsoft365Group': string;
    /** English native default: Security group · Entra */
    readonly securityGroup: string;
    /** English native default: Distribution list · Entra */
    readonly distributionList: string;
    /** English native default: Unverified email */
    readonly unvalidatedEmail: string;
  };
  /** Validation text. Every member is required in a complete dictionary. */
  readonly validation: {
    /** English native default: The selection limit has been reached. */
    readonly selectionLimitReached: string;
    /** English native default: The selection exceeds the configured limit. */
    readonly selectionLimitExceeded: string;
    /** English native default: Some selected values are unresolved. */
    readonly unresolvedValues: string;
    /** English native default: The selection contains values outside the configured scope. */
    readonly outOfScopeValues: string;
    /** English native default: Resolve the required IDs for the selected users. */
    readonly resolveRequiredUserIds: string;
  };
  /** Errors text. Every member is required in a complete dictionary. */
  readonly errors: {
    /** English native default: The candidate is outside the configured scope. */
    readonly candidateOutOfScope: string;
    /** English native default: The candidate could not be fully resolved. */
    readonly incompleteCandidateResolution: string;
    /** English native default: Resolution failed. */
    readonly resolutionFailed: string;
    /** English native default: Microsoft Graph client is unavailable. */
    readonly graphUnavailable: string;
  };
}

/** Optional instance dictionary, merged by known group and leaf. Undefined retains native defaults;
 * empty strings remain explicit overrides. Direct control labels/content/Field messages take precedence.
 * Changes update subsequent rendering and actions without resetting selection or changing hosted UI culture.
 * @example
 * ```ts
 * const strings: PickersStringsOverride = {
 *   status: { loading: 'Please wait' }, labels: { namedSearch: name => `Search ${name}` },
 * };
 * ```
 */
export interface PickersStringsOverride {
  /** Optional labels fields; omitted members retain native text. */
  readonly labels?: Partial<PickersStrings['labels']>;
  /** Optional status fields; omitted members retain native text. */
  readonly status?: Partial<PickersStrings['status']>;
  /** Optional sections fields; omitted members retain native text. */
  readonly sections?: Partial<PickersStrings['sections']>;
  /** Optional metadata fields; omitted members retain native text. */
  readonly metadata?: Partial<PickersStrings['metadata']>;
  /** Optional validation fields; omitted members retain native text. */
  readonly validation?: Partial<PickersStrings['validation']>;
  /** Optional errors fields; omitted members retain native text. */
  readonly errors?: Partial<PickersStrings['errors']>;
}
