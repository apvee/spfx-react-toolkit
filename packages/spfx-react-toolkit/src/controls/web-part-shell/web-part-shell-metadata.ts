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
 * @example Compose searchable title metadata in a web part
 * ```ts
 * import {
 *   IWebPartShellProperties,
 *   getWebPartShellPropertiesMetadata
 * } from '@apvee/spfx-react-toolkit';
 * import {
 *   BaseClientSideWebPart,
 *   IWebPartPropertiesMetadata
 * } from '@microsoft/sp-webpart-base';
 *
 * export class ExampleWebPart extends BaseClientSideWebPart<IWebPartShellProperties> {
 *   public render(): void {
 *     this.domElement.textContent = this.properties.title;
 *   }
 *   protected get propertiesMetadata(): IWebPartPropertiesMetadata {
 *     return { ...getWebPartShellPropertiesMetadata() };
 *   }
 * }
 * ```
 *
 * @scenario Inspect searchable Shell title metadata
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import { getWebPartShellPropertiesMetadata } from '@apvee/spfx-react-toolkit';
 * export function Demo({ resetRevision = 0 }: { resetRevision?: number } = {}): React.ReactElement {
 *   const [metadata, setMetadata] = React.useState(getWebPartShellPropertiesMetadata);
 *   const [fresh, setFresh] = React.useState<boolean | undefined>(undefined);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMetadata(getWebPartShellPropertiesMetadata());
 *     setFresh(undefined);
 *   }, [resetRevision]);
 *   return <div>
 *     <p>Plain title searchable: {metadata.title?.isSearchablePlainText ? 'yes' : 'no'}.</p>
 *     <p>titleHTML and showTitle are absent from this metadata map; they remain ordinary persisted properties.</p>
 *     <Button onClick={() => {
 *       const next = getWebPartShellPropertiesMetadata();
 *       setFresh(next !== metadata && next.title !== metadata.title);
 *       setMetadata(next);
 *     }}>Create fresh metadata</Button>
 *     <output aria-live="polite">{fresh === undefined ? 'No comparison requested.' : `Fresh map and title descriptor: ${fresh ? 'yes' : 'no'}.`}</output>
 *   </div>;
 * }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'The helper marks only the plain title as searchable. Create fresh metadata to compare independent maps and title descriptors.' };
 * ```
 *
 * @useWhen An SPFx web part persists shell title properties and must mark the plain title as searchable metadata.
 * @avoidWhen Keep titleHTML as ordinary data; do not mark rich serialized HTML as searchable plain text.
 */
export function getWebPartShellPropertiesMetadata(): IWebPartPropertiesMetadata {
  return { title: { isSearchablePlainText: true } };
}
