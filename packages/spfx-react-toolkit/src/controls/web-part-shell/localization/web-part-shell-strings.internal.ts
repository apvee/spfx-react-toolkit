/** Shell-owned title naming; common editor resources belong to EditableText. */
export interface WebPartShellStrings {
  readonly titleEditorLabel: string;
  readonly titlePlaceholder: string;
}

/** Construct immutable, instance-independent title naming resources. */
export function defineWebPartShellStrings(resource: WebPartShellStrings): WebPartShellStrings {
  return Object.freeze({ ...resource });
}
