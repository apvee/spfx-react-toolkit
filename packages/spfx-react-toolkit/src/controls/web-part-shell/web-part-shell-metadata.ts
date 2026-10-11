import type { IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';

/** Create fresh SPFx metadata marking only the plain title as searchable.
 * Compose the result with the web part's metadata; titleHTML remains ordinary data.
 * Does not import Property Pane runtime or mutate consumer metadata.
 * @returns A new map and title descriptor on every call.
 * @example
 * protected get propertiesMetadata(): IWebPartPropertiesMetadata {
 *   return { ...getWebPartShellPropertiesMetadata(), listId: { isSearchablePlainText: true } };
 * }
 */
export function getWebPartShellPropertiesMetadata(): IWebPartPropertiesMetadata {
  return { title: { isSearchablePlainText: true } };
}
