/** Private, data-only source contract; error keys preserve the stable internal error codes. */
export interface FileFolderPickerResource {
  readonly browse: {
    readonly files: string;
    readonly folders: string;
    readonly all: string;
  };
  readonly placeholders: {
    readonly files: string;
    readonly folders: string;
    readonly all: string;
  };
  readonly labels: {
    readonly selection: string;
    readonly file: string;
    readonly folder: string;
    readonly unresolvedItem: string;
    /** Template token: name. The selected item's caller/source text is inserted literally. */
    readonly removeItem: string;
  };
  readonly validation: {
    readonly invalidReferences: string;
    readonly controlledModeChanged: string;
    readonly incompatibleItems: string;
    readonly invalidPolicy: string;
    readonly selectionLimitExceeded: string;
  };
  /** Sanitized messages only: never insert token, claims, HTTP or native-error payloads. */
  readonly errors: {
    readonly configuration: string;
    readonly 'auth-resource': string;
    readonly 'auth-type': string;
    readonly 'auth-challenge': string;
    readonly 'auth-provider': string;
    readonly 'selection-identity': string;
    readonly 'selection-scope': string;
    readonly 'selection-policy': string;
    readonly 'selection-unavailable': string;
    readonly 'selection-invalid': string;
    readonly 'onedrive-unavailable': string;
    readonly opening: string;
    readonly timeout: string;
    readonly communication: string;
  };
  /** Toolkit-authored reply text; protocol discriminants and error codes are never translated. */
  readonly protocol: {
    readonly optionalAuthenticationTimeout: string;
    readonly unsupportedCommand: string;
    readonly optionalGraphAuthenticationFailed: string;
    readonly optionalPersonalAuthenticationFailed: string;
  };
}
