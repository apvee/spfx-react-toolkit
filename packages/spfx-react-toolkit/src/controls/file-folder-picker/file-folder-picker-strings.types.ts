/** Complete toolkit-owned FileFolderPicker UI dictionary. Native SPFx resources supply defaults.
 * Caller labels, content slots, values and backend errors keep their own text.
 * Callback results are plain text; source arguments are passed literally without template reprocessing.
 * @example
 * ```ts
 * function customize(strings: FileFolderPickerStrings): FileFolderPickerStrings {
 *   return { ...strings, labels: { ...strings.labels, selection: 'Attachments' } };
 * }
 * ```
 */
export interface FileFolderPickerStrings {
  /** Browse text. Every member is required in a complete dictionary. */
  readonly browse: {
    /** English native default: Browse files… */
    readonly files: string;
    /** English native default: Browse folders… */
    readonly folders: string;
    /** English native default: Browse files and folders… */
    readonly all: string;
  };
  /** Placeholders text. Every member is required in a complete dictionary. */
  readonly placeholders: {
    /** English native default: Select files… */
    readonly files: string;
    /** English native default: Select folders… */
    readonly folders: string;
    /** English native default: Select files or folders… */
    readonly all: string;
  };
  /** Labels text. Every member is required in a complete dictionary. */
  readonly labels: {
    /** English native default: File and folder selection */
    readonly selection: string;
    /** English native default: File */
    readonly file: string;
    /** English native default: Folder */
    readonly folder: string;
    /** English native default: Unresolved item */
    readonly unresolvedItem: string;
    /** Selected item display name, inserted literally.
     * @param name - Caller/source selected display name.
     * @returns Plain accessible remove-action text.
     */
    readonly removeItem: (name: string) => string;
  };
  /** Validation text. Every member is required in a complete dictionary. */
  readonly validation: {
    /** English native default: Invalid file/folder references. */
    readonly invalidReferences: string;
    /** English native default: Controlled mode cannot change while the control is mounted. */
    readonly controlledModeChanged: string;
    /** English native default: The selection contains items incompatible with the configured type, category or extensions. Remove or replace them. */
    readonly incompatibleItems: string;
    /** English native default: Invalid selection type, category, extensions or limit configuration. */
    readonly invalidPolicy: string;
    /** English native default: The selection exceeds the configured limit. Remove the extra items. */
    readonly selectionLimitExceeded: string;
  };
  /** Errors text. Every member is required in a complete dictionary. */
  readonly errors: {
    /** English native default: Invalid picker configuration. Check the source, types, category and limits; use the preset without legacy options. */
    readonly configuration: string;
    /** English native default: The picker requested an authentication resource that is not allowed by the configuration. */
    readonly 'auth-resource': string;
    /** English native default: The picker requested an authentication type that is not supported by the configuration. */
    readonly 'auth-type': string;
    /** English native default: The picker requires an additional authentication check that is not supported. */
    readonly 'auth-challenge': string;
    /** English native default: Could not obtain the SharePoint token. Try again; if the problem persists, check your access and existing permissions. */
    readonly 'auth-provider': string;
    /** English native default: Could not verify the selected item's identity. Reopen the picker and try again. */
    readonly 'selection-identity': string;
    /** English native default: The selected item is outside the allowed site, library or folder. */
    readonly 'selection-scope': string;
    /** English native default: The selection does not meet the allowed type, category, extensions or item count. */
    readonly 'selection-policy': string;
    /** English native default: The selected item is unavailable or inaccessible. */
    readonly 'selection-unavailable': string;
    /** English native default: Could not verify the selection. No new items were added. */
    readonly 'selection-invalid': string;
    /** English native default: The current user's OneDrive is unavailable or inaccessible. Check that the drive exists and try again. */
    readonly 'onedrive-unavailable': string;
    /** English native default: Could not open the picker. Check the configuration and access to the source, then try again. */
    readonly opening: string;
    /** English native default: The picker did not respond in time. Close it and try again. */
    readonly timeout: string;
    /** English native default: Communication with the picker was interrupted. Reopen it and try again. */
    readonly communication: string;
  };
  /** Protocol text. Every member is required in a complete dictionary. */
  readonly protocol: {
    /** English native default: The optional authentication feature did not respond in time. Try again. */
    readonly optionalAuthenticationTimeout: string;
    /** English native default: This integration does not support that command. */
    readonly unsupportedCommand: string;
    /** English native default: Could not authenticate the optional Microsoft Graph feature. */
    readonly optionalGraphAuthenticationFailed: string;
    /** English native default: Could not authenticate the optional personal destination. */
    readonly optionalPersonalAuthenticationFailed: string;
  };
}

/** Optional instance dictionary, merged by known group and leaf. Undefined retains native defaults;
 * empty strings remain explicit overrides. Direct control labels/content/Field messages take precedence.
 * Changes update subsequent rendering and actions without resetting selection or changing hosted UI culture.
 * @example
 * ```ts
 * const strings: FileFolderPickerStringsOverride = {
 *   browse: { files: 'Choose documents' }, errors: { timeout: 'Please reopen the picker.' },
 * };
 * ```
 */
export interface FileFolderPickerStringsOverride {
  /** Optional browse fields; omitted members retain native text. */
  readonly browse?: Partial<FileFolderPickerStrings['browse']>;
  /** Optional placeholders fields; omitted members retain native text. */
  readonly placeholders?: Partial<FileFolderPickerStrings['placeholders']>;
  /** Optional labels fields; omitted members retain native text. */
  readonly labels?: Partial<FileFolderPickerStrings['labels']>;
  /** Optional validation fields; omitted members retain native text. */
  readonly validation?: Partial<FileFolderPickerStrings['validation']>;
  /** Optional errors fields; omitted members retain native text. */
  readonly errors?: Partial<FileFolderPickerStrings['errors']>;
  /** Optional protocol fields; omitted members retain native text. */
  readonly protocol?: Partial<FileFolderPickerStrings['protocol']>;
}
