/**
 * Actual SPFx web-part integration, persisted editable titles and inherited host presentation.
 *
 * @packageDocumentation
 * @docPage controls/web-part-shell
 */
import * as React from 'react';
import { SPFxWebPartProvider } from '../../core/provider-webpart';
import { WebPartShellRuntime } from './web-part-shell-runtime.internal';
import type { IWebPartShellProperties, WebPartShellProps } from './web-part-shell.types';

function WebPartShellRender<TProperties extends IWebPartShellProperties>(
  { instance, ...props }: WebPartShellProps<TProperties>, ref: React.ForwardedRef<HTMLDivElement>
): React.ReactElement {
  return <SPFxWebPartProvider instance={instance}>
    <WebPartShellRuntime {...props} key={instance.instanceId} ownerDocument={instance.context.domElement.ownerDocument} rootRef={ref} />
  </SPFxWebPartProvider>;
}

// Explicit base instantiation below prevents slot-union expansion during React 17 inference;
// it preserves the public generic instance contract and the original render function.
/**
 * Compose SPFx and Fluent providers for an actual web part and its stable content.
 * Same-instance rerenders, title Read/Edit, visibility/placement, Card size and non-subtle appearance
 * updates retain children; a new instanceId resets them. Subtle/Card transitions can remount children.
 * The existing provider waits for ServiceScope readiness and isolates runtime state.
 * Theme precedence is explicit override, ancestor Fluent, live host, then light fallback.
 * Uses the web part ownerDocument and inherited Griffel renderer; hosts must supply a
 * renderer for that document. Subtle (default) adds no Card. Filled, filled-alternative
 * and outline render a native vertical Card with focusMode off; size defaults to medium. Title placement
 * defaults to outside; inside places the complete header into Card. Subtle ignores placement and size.
 * Section adds no outer inset; application auto padding
 * is medium below a 640px container inline size and large from 640px.
 * Shell owns outer padding and header/body spacing; native Card owns its inner spacing.
 * Title and header actions are vertically centered within their wrapping row.
 * Hiding the title removes the complete header and its gap without evaluating action callbacks.
 * Header placement changes can remount the header. Appearance/size/titlePosition are optional React
 * options; no additional persisted properties are required.
 * Native div attributes/events, className, caller style and ref reach the shell root.
 * No normalized/default title property is written at mount.
 * @param props - Web part instance, surface/theme options and native div attributes.
 * @returns The provider-composed shell, once services are ready.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { IWebPartShellProperties, WebPartShell } from '@apvee/spfx-react-toolkit';
 * import { Button } from '@fluentui/react-components';
 * import * as ReactDom from 'react-dom';
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * export class ExampleWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     ReactDom.render(<WebPartShell instance={this} appearance="filled" size="small" titlePosition="inside"
 *       action={<Button onClick={() => this.render()}>Refresh</Button>}>
 *       <p>Workspace content</p>
 *     </WebPartShell>, this.domElement);
 *   }
 *   protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
 * }
 * ```
 *
 * @typeParam TProperties - Persisted web-part property bag extending the shell’s title properties.
 * @useWhen A real SPFx web-part instance needs matching providers, host spacing and persisted editable title integration.
 * @avoidWhen Use the matching direct SPFx provider for hosts that are not web parts.
 *
 *
 * @example First web part with metadata and native Property Pane
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { BaseClientSideWebPart, type IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
 * import { PropertyPaneToggle, type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import {
 *   WebPartShell,
 *   getWebPartShellPropertiesMetadata,
 *   useSPFxProperties,
 *   useSPFxUserInfo,
 *   type IWebPartShellProperties,
 * } from '@apvee/spfx-react-toolkit';
 *
 * interface IExampleWebPartProperties extends IWebPartShellProperties {
 *   description: string;
 * }
 *
 * function Content(): React.ReactElement {
 *   const { displayName } = useSPFxUserInfo();
 *   const { properties } = useSPFxProperties<IExampleWebPartProperties>();
 *   return <p>Hello {displayName}! {properties?.description}</p>;
 * }
 *
 * export default class ExampleWebPart extends BaseClientSideWebPart<IExampleWebPartProperties> {
 *   public render(): void {
 *     ReactDom.render(
 *       <WebPartShell instance={this} appearance="outline" size="medium" titlePosition="outside">
 *         <Content />
 *       </WebPartShell>,
 *       this.domElement
 *     );
 *   }
 *
 *   protected get propertiesMetadata(): IWebPartPropertiesMetadata {
 *     return {
 *       ...getWebPartShellPropertiesMetadata(),
 *       description: { isSearchablePlainText: true },
 *     };
 *   }
 *
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return { pages: [{ groups: [{ groupName: 'Title', groupFields: [
 *       PropertyPaneToggle('showTitle', { label: 'Show title', checked: this.properties.showTitle ?? true }),
 *     ] }] }] };
 *   }
 *
 *   protected onDispose(): void {
 *     ReactDom.unmountComponentAtNode(this.domElement);
 *   }
 * }
 * ```
 *
 *
 * @example Customize the header action and restore a hidden title
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import { WebPartShell, useSPFxProperties, type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
 * function HideTitle() {
 *   const { setProperties } = useSPFxProperties<IWebPartShellProperties>();
 *   return <button type="button" onClick={() => setProperties({ showTitle: false })}>Hide title</button>;
 * }
 * function ShowTitle() {
 *   const { setProperties } = useSPFxProperties<IWebPartShellProperties>();
 *   return <button type="button" onClick={() => setProperties({ showTitle: true })}>Show title</button>;
 * }
 * export default class HeaderActionWebPart extends BaseClientSideWebPart<IWebPartShellProperties> {
 *   private readonly actionRef = React.createRef<HTMLDivElement>();
 *   public render(): void {
 *     ReactDom.render(<WebPartShell instance={this} action={{ ref: this.actionRef,
 *       children: (Component, props) => <Component {...props}><HideTitle /></Component>,
 *     }}><ShowTitle /><p>Body state belongs to the application.</p></WebPartShell>, this.domElement);
 *   }
 *   protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
 * }
 * ```
 *
 * @example Configure the actual web part Shell
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select } from '@fluentui/react-components';
 * import { useSPFxProperties, useSPFxDisplayMode, type IWebPartShellProperties, type WebPartShellProps } from '@apvee/spfx-react-toolkit';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * interface DemoShellPort {
 *   presentation: Required<Pick<WebPartShellProps<IWebPartShellProperties>, 'appearance' | 'size' | 'titlePosition' | 'padding'>>;
 *   updatePresentation(patch: Partial<DemoShellPort['presentation']>): void;
 * }
 * interface DemoProps {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 *   shell?: DemoShellPort;
 * }
 *
 * // The host supplies its existing Shell and initialized SPFx provider, never a nested Shell.
 * export function Demo({ render = standalone, resetRevision = 0, shell }: DemoProps = {}): React.ReactElement {
 *   return shell === undefined
 *     ? <p>This scenario requires the actual WebPartShell host, its presentation port and initialized SPFx provider.</p>
 *     : <ShellContent render={render} resetRevision={resetRevision} shell={shell} />;
 * }
 * function ShellContent({ render, resetRevision, shell }: {
 *   render: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision: number;
 *   shell: DemoShellPort;
 * }): React.ReactElement {
 *   const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
 *   const { isEdit } = useSPFxDisplayMode();
 *   const [note, setNote] = React.useState('Intranet project review');
 *   const [count, setCount] = React.useState(0);
 *   const showButton = React.useRef<HTMLButtonElement>(null); // React 17 ref interoperability requires null.
 *   const revision = React.useRef(resetRevision);
 *   const updateBaseline = React.useCallback((): void => {
 *     const baseline: DemoShellPort['presentation'] = { appearance: 'subtle', size: 'medium', titlePosition: 'outside', padding: 'auto' };
 *     const current = shell.presentation;
 *     if (current.appearance !== baseline.appearance || current.size !== baseline.size ||
 *         current.titlePosition !== baseline.titlePosition || current.padding !== baseline.padding) {
 *       shell.updatePresentation(baseline);
 *     }
 *   }, [shell]);
 *   React.useEffect(() => {
 *     // Mount and native Read/Edit updates only observe the revision. A later
 *     // revision is an explicit reset; no persisted properties are written on mount.
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setNote('Intranet project review');
 *     setCount(0);
 *     updateBaseline();
 *     if (isEdit) setProperties({ title: 'SPFx React Toolkit', titleHTML: undefined, showTitle: true });
 *   }, [resetRevision, updateBaseline, isEdit, setProperties]);
 *   React.useEffect(() => {
 *     if (properties?.showTitle === false) showButton.current?.focus();
 *   }, [properties?.showTitle]);
 *   return render({
 *     controls: <>
 *       <Field label="Appearance"><Select aria-label="Appearance" value={shell.presentation.appearance}
 *         onChange={event => shell.updatePresentation({ appearance: event.target.value === 'filled' ? 'filled' : event.target.value === 'filled-alternative' ? 'filled-alternative' : event.target.value === 'outline' ? 'outline' : 'subtle' })}>
 *         <option value="subtle">Subtle</option><option value="filled">Filled</option><option value="filled-alternative">Filled alternative</option><option value="outline">Outline</option>
 *       </Select></Field>
 *       <Field label="Card size"><Select aria-label="Card size" value={shell.presentation.size}
 *         onChange={event => shell.updatePresentation({ size: event.target.value === 'small' ? 'small' : event.target.value === 'large' ? 'large' : 'medium' })}>
 *         <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *       </Select></Field>
 *       <Field label="Title position"><Select aria-label="Title position" value={shell.presentation.titlePosition}
 *         onChange={event => shell.updatePresentation({ titlePosition: event.target.value === 'inside' ? 'inside' : 'outside' })}>
 *         <option value="outside">Outside</option><option value="inside">Inside</option>
 *       </Select></Field>
 *       <Field label="Spacing"><Select aria-label="Spacing" value={shell.presentation.padding}
 *         onChange={event => shell.updatePresentation({ padding: event.target.value === 'none' ? 'none' : event.target.value === 'small' ? 'small' : event.target.value === 'medium' ? 'medium' : event.target.value === 'large' ? 'large' : 'auto' })}>
 *         <option value="auto">Auto</option><option value="none">None</option><option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *       </Select></Field>
 *     </>,
 *     commands: <>
 *       <Button onClick={() => shell.updatePresentation({ appearance: 'outline' })}>Use Card baseline</Button>
 *       <Button onClick={() => shell.updatePresentation({ appearance: 'subtle' })}>Use subtle baseline</Button>
 *       <Button onClick={() => setProperties({ showTitle: false })}>Hide title</Button>
 *       <Button ref={showButton} onClick={() => setProperties({ showTitle: true })}>Show title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: 'Project documents', titleHTML: undefined }); }}>Plain title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: 'Project documents', titleHTML: '<h2><strong>Project</strong> documents</h2>' }); }}>Formatted title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: '', titleHTML: '' }); }}>Empty title</Button>
 *     </>,
 *     status: <div role="status" aria-live="polite"><p>Mode: {isEdit ? 'Edit' : 'Read'} · Child counter: {count}</p>
 *       <p>Use the actual host title and header action. Presentation presets are explicit; Reset restores this source baseline and restores the title only in Edit.</p></div>,
 *     children: <>
 *       <Input aria-label="Project note" value={note} onChange={(_event, data) => setNote(data.value)} />
 *       <Button onClick={() => setCount(current => current + 1)}>Increment counter</Button>
 *     </>,
 *     diagnostics: <output>{JSON.stringify({ title: properties?.title, titleHTML: properties?.titleHTML, showTitle: properties?.showTitle, ...shell.presentation })}</output>,
 *   });
 * }
 * export const demo = {
 *   host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', hostAdapter: 'webPartShell',
 *   observe: 'Configure the actual outer Shell. Its title and action are outside this scenario; native Read/Edit comes from SharePoint. Reset writes the source title only in Edit.',
 * };
 * ```
 *
 * @example Retain content through Shell presentation changes
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select } from '@fluentui/react-components';
 * import { useSPFxProperties, useSPFxDisplayMode, type IWebPartShellProperties, type WebPartShellProps } from '@apvee/spfx-react-toolkit';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * interface DemoShellPort {
 *   presentation: Required<Pick<WebPartShellProps<IWebPartShellProperties>, 'appearance' | 'size' | 'titlePosition' | 'padding'>>;
 *   updatePresentation(patch: Partial<DemoShellPort['presentation']>): void;
 * }
 * interface DemoProps {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 *   shell?: DemoShellPort;
 * }
 *
 * // The host supplies its existing Shell and initialized SPFx provider, never a nested Shell.
 * export function Demo({ render = standalone, resetRevision = 0, shell }: DemoProps = {}): React.ReactElement {
 *   return shell === undefined
 *     ? <p>This scenario requires the actual WebPartShell host, its presentation port and initialized SPFx provider.</p>
 *     : <ShellContent render={render} resetRevision={resetRevision} shell={shell} />;
 * }
 * function ShellContent({ render, resetRevision, shell }: {
 *   render: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision: number;
 *   shell: DemoShellPort;
 * }): React.ReactElement {
 *   const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
 *   const { isEdit } = useSPFxDisplayMode();
 *   const [note, setNote] = React.useState('Intranet project review');
 *   const [count, setCount] = React.useState(0);
 *   const showButton = React.useRef<HTMLButtonElement>(null); // React 17 ref interoperability requires null.
 *   const revision = React.useRef(resetRevision);
 *   const updateBaseline = React.useCallback((): void => {
 *     const baseline: DemoShellPort['presentation'] = { appearance: 'outline', size: 'medium', titlePosition: 'outside', padding: 'auto' };
 *     const current = shell.presentation;
 *     if (current.appearance !== baseline.appearance || current.size !== baseline.size ||
 *         current.titlePosition !== baseline.titlePosition || current.padding !== baseline.padding) {
 *       shell.updatePresentation(baseline);
 *     }
 *   }, [shell]);
 *   React.useEffect(() => {
 *     // Mount and native Read/Edit updates only observe the revision. A later
 *     // revision is an explicit reset; no persisted properties are written on mount.
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setNote('Intranet project review');
 *     setCount(0);
 *     updateBaseline();
 *     if (isEdit) setProperties({ title: 'SPFx React Toolkit', titleHTML: undefined, showTitle: true });
 *   }, [resetRevision, updateBaseline, isEdit, setProperties]);
 *   React.useEffect(() => {
 *     if (properties?.showTitle === false) showButton.current?.focus();
 *   }, [properties?.showTitle]);
 *   return render({
 *     controls: <>
 *       <Field label="Appearance"><Select aria-label="Appearance" value={shell.presentation.appearance}
 *         onChange={event => shell.updatePresentation({ appearance: event.target.value === 'filled' ? 'filled' : event.target.value === 'filled-alternative' ? 'filled-alternative' : event.target.value === 'outline' ? 'outline' : 'subtle' })}>
 *         <option value="subtle">Subtle</option><option value="filled">Filled</option><option value="filled-alternative">Filled alternative</option><option value="outline">Outline</option>
 *       </Select></Field>
 *       <Field label="Card size"><Select aria-label="Card size" value={shell.presentation.size}
 *         onChange={event => shell.updatePresentation({ size: event.target.value === 'small' ? 'small' : event.target.value === 'large' ? 'large' : 'medium' })}>
 *         <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *       </Select></Field>
 *       <Field label="Title position"><Select aria-label="Title position" value={shell.presentation.titlePosition}
 *         onChange={event => shell.updatePresentation({ titlePosition: event.target.value === 'inside' ? 'inside' : 'outside' })}>
 *         <option value="outside">Outside</option><option value="inside">Inside</option>
 *       </Select></Field>
 *       <Field label="Spacing"><Select aria-label="Spacing" value={shell.presentation.padding}
 *         onChange={event => shell.updatePresentation({ padding: event.target.value === 'none' ? 'none' : event.target.value === 'small' ? 'small' : event.target.value === 'medium' ? 'medium' : event.target.value === 'large' ? 'large' : 'auto' })}>
 *         <option value="auto">Auto</option><option value="none">None</option><option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *       </Select></Field>
 *     </>,
 *     commands: <>
 *       <Button onClick={() => shell.updatePresentation({ appearance: 'outline' })}>Use Card baseline</Button>
 *       <Button onClick={() => shell.updatePresentation({ appearance: 'subtle' })}>Use subtle baseline</Button>
 *       <Button onClick={() => setProperties({ showTitle: false })}>Hide title</Button>
 *       <Button ref={showButton} onClick={() => setProperties({ showTitle: true })}>Show title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: 'Project documents', titleHTML: undefined }); }}>Plain title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: 'Project documents', titleHTML: '<h2><strong>Project</strong> documents</h2>' }); }}>Formatted title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: '', titleHTML: '' }); }}>Empty title</Button>
 *     </>,
 *     status: <div role="status" aria-live="polite"><p>Mode: {isEdit ? 'Edit' : 'Read'} · Child counter: {count}</p>
 *       <p>Choose Use Card baseline first, then edit the note and counter. Card appearance, size, title position and visibility preserve content.</p></div>,
 *     children: <>
 *       <Input aria-label="Project note" value={note} onChange={(_event, data) => setNote(data.value)} />
 *       <Button onClick={() => setCount(current => current + 1)}>Increment counter</Button>
 *     </>,
 *     diagnostics: <output>{JSON.stringify({ title: properties?.title, titleHTML: properties?.titleHTML, showTitle: properties?.showTitle, ...shell.presentation })}</output>,
 *   });
 * }
 * export const demo = {
 *   host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', hostAdapter: 'webPartShell',
 *   observe: 'Choose Use Card baseline, edit the note and increment. Change between Card variants, size, title position and visibility: the same content remains mounted.',
 * };
 * ```
 *
 * @example Observe content remount between subtle and Card
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select } from '@fluentui/react-components';
 * import { useSPFxProperties, useSPFxDisplayMode, type IWebPartShellProperties, type WebPartShellProps } from '@apvee/spfx-react-toolkit';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * interface DemoShellPort {
 *   presentation: Required<Pick<WebPartShellProps<IWebPartShellProperties>, 'appearance' | 'size' | 'titlePosition' | 'padding'>>;
 *   updatePresentation(patch: Partial<DemoShellPort['presentation']>): void;
 * }
 * interface DemoProps {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 *   shell?: DemoShellPort;
 * }
 *
 * // The host supplies its existing Shell and initialized SPFx provider, never a nested Shell.
 * export function Demo({ render = standalone, resetRevision = 0, shell }: DemoProps = {}): React.ReactElement {
 *   return shell === undefined
 *     ? <p>This scenario requires the actual WebPartShell host, its presentation port and initialized SPFx provider.</p>
 *     : <ShellContent render={render} resetRevision={resetRevision} shell={shell} />;
 * }
 * function ShellContent({ render, resetRevision, shell }: {
 *   render: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision: number;
 *   shell: DemoShellPort;
 * }): React.ReactElement {
 *   const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
 *   const { isEdit } = useSPFxDisplayMode();
 *   const [note, setNote] = React.useState('Intranet project review');
 *   const [count, setCount] = React.useState(0);
 *   const showButton = React.useRef<HTMLButtonElement>(null); // React 17 ref interoperability requires null.
 *   const revision = React.useRef(resetRevision);
 *   const updateBaseline = React.useCallback((): void => {
 *     const baseline: DemoShellPort['presentation'] = { appearance: 'subtle', size: 'medium', titlePosition: 'outside', padding: 'auto' };
 *     const current = shell.presentation;
 *     if (current.appearance !== baseline.appearance || current.size !== baseline.size ||
 *         current.titlePosition !== baseline.titlePosition || current.padding !== baseline.padding) {
 *       shell.updatePresentation(baseline);
 *     }
 *   }, [shell]);
 *   React.useEffect(() => {
 *     // Mount and native Read/Edit updates only observe the revision. A later
 *     // revision is an explicit reset; no persisted properties are written on mount.
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setNote('Intranet project review');
 *     setCount(0);
 *     updateBaseline();
 *     if (isEdit) setProperties({ title: 'SPFx React Toolkit', titleHTML: undefined, showTitle: true });
 *   }, [resetRevision, updateBaseline, isEdit, setProperties]);
 *   React.useEffect(() => {
 *     if (properties?.showTitle === false) showButton.current?.focus();
 *   }, [properties?.showTitle]);
 *   return render({
 *     controls: <>
 *       <Field label="Appearance"><Select aria-label="Appearance" value={shell.presentation.appearance}
 *         onChange={event => shell.updatePresentation({ appearance: event.target.value === 'filled' ? 'filled' : event.target.value === 'filled-alternative' ? 'filled-alternative' : event.target.value === 'outline' ? 'outline' : 'subtle' })}>
 *         <option value="subtle">Subtle</option><option value="filled">Filled</option><option value="filled-alternative">Filled alternative</option><option value="outline">Outline</option>
 *       </Select></Field>
 *       <Field label="Card size"><Select aria-label="Card size" value={shell.presentation.size}
 *         onChange={event => shell.updatePresentation({ size: event.target.value === 'small' ? 'small' : event.target.value === 'large' ? 'large' : 'medium' })}>
 *         <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *       </Select></Field>
 *       <Field label="Title position"><Select aria-label="Title position" value={shell.presentation.titlePosition}
 *         onChange={event => shell.updatePresentation({ titlePosition: event.target.value === 'inside' ? 'inside' : 'outside' })}>
 *         <option value="outside">Outside</option><option value="inside">Inside</option>
 *       </Select></Field>
 *       <Field label="Spacing"><Select aria-label="Spacing" value={shell.presentation.padding}
 *         onChange={event => shell.updatePresentation({ padding: event.target.value === 'none' ? 'none' : event.target.value === 'small' ? 'small' : event.target.value === 'medium' ? 'medium' : event.target.value === 'large' ? 'large' : 'auto' })}>
 *         <option value="auto">Auto</option><option value="none">None</option><option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *       </Select></Field>
 *     </>,
 *     commands: <>
 *       <Button onClick={() => shell.updatePresentation({ appearance: 'outline' })}>Use Card baseline</Button>
 *       <Button onClick={() => shell.updatePresentation({ appearance: 'subtle' })}>Use subtle baseline</Button>
 *       <Button onClick={() => setProperties({ showTitle: false })}>Hide title</Button>
 *       <Button ref={showButton} onClick={() => setProperties({ showTitle: true })}>Show title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: 'Project documents', titleHTML: undefined }); }}>Plain title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: 'Project documents', titleHTML: '<h2><strong>Project</strong> documents</h2>' }); }}>Formatted title</Button>
 *       <Button disabled={!isEdit} onClick={() => { if (isEdit) setProperties({ title: '', titleHTML: '' }); }}>Empty title</Button>
 *     </>,
 *     status: <div role="status" aria-live="polite"><p>Mode: {isEdit ? 'Edit' : 'Read'} · Child counter: {count}</p>
 *       <p>Edit the note and increment, then switch subtle and Card. The real Shell may remount its content; upper explorer selection survives.</p></div>,
 *     children: <>
 *       <Input aria-label="Project note" value={note} onChange={(_event, data) => setNote(data.value)} />
 *       <Button onClick={() => setCount(current => current + 1)}>Increment counter</Button>
 *     </>,
 *     diagnostics: <output>{JSON.stringify({ title: properties?.title, titleHTML: properties?.titleHTML, showTitle: properties?.showTitle, ...shell.presentation })}</output>,
 *   });
 * }
 * export const demo = {
 *   host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', hostAdapter: 'webPartShell',
 *   observe: 'Edit the note and counter, then switch subtle and Card with the explicit commands. Child state resets at the real body boundary; group, feature and scenario selection survive.',
 * };
 * ```
 *

 */
export const WebPartShell = React.forwardRef<HTMLDivElement, WebPartShellProps<IWebPartShellProperties>>(WebPartShellRender<IWebPartShellProperties>) as <TProperties extends IWebPartShellProperties>(
  // instance is currently the only TProperties-dependent member; keep any future generic
  // members in the inferred portion. Comparing a generic full-props intersection makes
  // React 17 JSX expand the Fluent content-slot union for props spreads.
  props: Pick<WebPartShellProps<TProperties>, 'instance'> &
    Omit<WebPartShellProps<IWebPartShellProperties>, 'instance'> & React.RefAttributes<HTMLDivElement>
) => React.ReactElement;
