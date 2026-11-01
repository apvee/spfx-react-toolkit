/**
 * @packageDocumentation
 * @docPage controls/web-part-shell
 */
import type { IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';

/**
 * Create fresh SPFx metadata marking only the plain title as searchable.
 * Compose the result with the web part's metadata; titleHTML remains ordinary data.
 * Does not import Property Pane runtime or mutate consumer metadata.
 * @returns A new map and title descriptor on every call.
 * @example
 * ```ts
 * import { IWebPartShellProperties, getWebPartShellPropertiesMetadata } from '@apvee/spfx-react-toolkit';
 * import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
 * export class ExampleWebPart extends BaseClientSideWebPart<IWebPartShellProperties> {
 *   public render(): void { this.domElement.textContent = this.properties.title; }
 *   protected get propertiesMetadata(): IWebPartPropertiesMetadata {
 *     return { ...getWebPartShellPropertiesMetadata() };
 *   }
 * }
 * ```
 *
 * @useWhen An SPFx web part persists shell title properties and must mark the plain title as searchable metadata.
 * @avoidWhen Keep titleHTML as ordinary data; do not mark rich serialized HTML as searchable plain text.
 */
export function getWebPartShellPropertiesMetadata(): IWebPartPropertiesMetadata {
  return { title: { isSearchablePlainText: true } };
}
