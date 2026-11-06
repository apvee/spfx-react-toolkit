/**
 * Mount a native Application Customizer React surface with an isolated toolkit runtime.
 *
 * @packageDocumentation
 * @docPage core/providers
 */
// provider-application-customizer.tsx
// Type-safe provider for Application Customizers

import * as React from 'react';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Initialized native host instance and optional subtree for SPFxApplicationCustomizerProvider. Each mounted provider owns a separate runtime store. Children mount after the active ServiceScope finishes; missing-scope fallback does not make native services available.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @example Typed native host input
 * ```tsx
 * import * as React from 'react';
 * import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
 * import { SPFxApplicationCustomizerProvider, SPFxApplicationCustomizerProviderProps } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(instance: BaseApplicationCustomizer<{ title: string }>): React.ReactElement {
 *   const props: SPFxApplicationCustomizerProviderProps<{ title: string }> = { instance, children: <span>Host content</span> };
 *   return <SPFxApplicationCustomizerProvider {...props} />;
 * }
 * ```
 */
export interface SPFxApplicationCustomizerProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized Application Customizer instance supplying context, properties and lifecycle integration.
   */
  instance: BaseApplicationCustomizer<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omitted or undefined children render no application content.
   */
  children?: React.ReactNode;
}

/**
 * Establishes SPFx context and isolated runtime state for an Application Customizer React surface. Pass its initialized native instance. Descendants mount after the active ServiceScope finishes; the provider never registers, finishes or authorizes services. The missing-scope fallback allows rendering but service-consuming hooks can fail.
 *
 * Host renders reconcile shallow property snapshots, including removals and reference changes; nested in-place changes are not deep-observed. Hook property updates synchronize back to the native instance. Theme subscriptions and runtime listeners are cleaned up on unmount; obsolete scope-readiness callbacks are ignored. Keep a mounted provider associated with its owning host surface and unmount React when that surface is disposed.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @param props - Required initialized BaseApplicationCustomizer instance and optional children; omitted or undefined children render no application content.
 * @returns Provider element establishing context after ServiceScope readiness.
 * @useWhen toolkit hooks render inside a native ApplicationCustomizer surface owned by this host instance.
 * @avoidWhen the host is another SPFx component kind; use its matching typed provider.
 * @example Mount an initialized native host
 * ```tsx
 * import * as React from 'react';
 * import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
 * import {
 *   SPFxApplicationCustomizerProvider,
 *   SPFxApplicationCustomizerProviderProps
 * } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(
 *   instance: BaseApplicationCustomizer<{ title: string }>
 * ): React.ReactElement {
 *   const props: SPFxApplicationCustomizerProviderProps<{ title: string }> = {
 *     instance,
 *     children: <span>Host content</span>
 *   };
 *   return <SPFxApplicationCustomizerProvider {...props} />;
 * }
 * ```
 * @public
 * @docPage core/providers
 *
 *
 * @example Retry late placeholders and dispose captured roots
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
 * import { SPFxApplicationCustomizerProvider, useSPFxProperties } from '@apvee/spfx-react-toolkit';
 *
 * interface IMyCustomizerProps {
 *   headerMessage: string;
 * }
 *
 * const HeaderComponent: React.FC = () => {
 *   const { properties } = useSPFxProperties<IMyCustomizerProps>();
 *   return <header>{properties?.headerMessage ?? 'Welcome'}</header>;
 * };
 *
 * export default class MyApplicationCustomizer extends BaseApplicationCustomizer<IMyCustomizerProps> {
 *   private topPlaceholder?: PlaceholderContent;
 *
 *   public onInit(): Promise<void> {
 *     this.context.placeholderProvider.changedEvent.add(this, this.renderTopPlaceholder);
 *     this.renderTopPlaceholder();
 *     return Promise.resolve();
 *   }
 *
 *   public onDispose(): void {
 *     this.context.placeholderProvider.changedEvent.remove(this, this.renderTopPlaceholder);
 *     const placeholder = this.topPlaceholder;
 *     this.topPlaceholder = undefined;
 *     if (placeholder) {
 *       ReactDom.unmountComponentAtNode(placeholder.domElement);
 *       placeholder.dispose();
 *     }
 *   }
 *
 *   private renderTopPlaceholder = (): void => {
 *     if (!this.topPlaceholder) {
 *       const capturedPlaceholder = this.context.placeholderProvider.tryCreateContent(
 *         PlaceholderName.Top,
 *         {
 *           onDispose: () => {
 *             if (capturedPlaceholder) {
 *               ReactDom.unmountComponentAtNode(capturedPlaceholder.domElement);
 *               if (this.topPlaceholder === capturedPlaceholder) {
 *                 this.topPlaceholder = undefined;
 *               }
 *             }
 *           }
 *         }
 *       );
 *       this.topPlaceholder = capturedPlaceholder;
 *     }
 *
 *     if (!this.topPlaceholder) return;
 *     const element = React.createElement(
 *       SPFxApplicationCustomizerProvider,
 *       { instance: this },
 *       React.createElement(HeaderComponent)
 *     );
 *     ReactDom.render(element, this.topPlaceholder.domElement);
 *   };
 * }
 * ```

 * @scenario Compare Application Customizer provider availability with the current host
 * ```tsx
 * import * as React from 'react';
 * import { Accordion, AccordionHeader, AccordionItem, AccordionPanel, Badge, Field, Select, makeStyles, tokens } from '@fluentui/react-components';
 * import { SPFxWebPartProvider, SPFxApplicationCustomizerProvider, SPFxFieldCustomizerProvider, SPFxListViewCommandSetProvider, useSPFxContext } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ result: { display: 'grid', rowGap: tokens.spacingVerticalS }, facts: { display: 'grid', gridTemplateColumns: 'auto 1fr', columnGap: tokens.spacingHorizontalM } });
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const styles = useStyles();
 *   const context = useSPFxContext();
 *   const providers = [
 *     { label: 'SPFxWebPartProvider', value: SPFxWebPartProvider, kind: 'WebPart' },
 *     { label: 'SPFxApplicationCustomizerProvider', value: SPFxApplicationCustomizerProvider, kind: 'AppCustomizer' },
 *     { label: 'SPFxFieldCustomizerProvider', value: SPFxFieldCustomizerProvider, kind: 'FieldCustomizer' },
 *     { label: 'SPFxListViewCommandSetProvider', value: SPFxListViewCommandSetProvider, kind: 'CommandSet' },
 *   ];
 *   const [selection, setSelection] = React.useState('SPFxApplicationCustomizerProvider');
 *   const revision = React.useRef(resetRevision);
 *   const lifetime = React.useMemo(() => ({ active: true }), [resetRevision]);
 *   React.useLayoutEffect(() => {
 *     if (revision.current !== resetRevision) {
 *       revision.current = resetRevision;
 *       setSelection('SPFxApplicationCustomizerProvider');
 *     }
 *     return () => { lifetime.active = false; };
 *   }, [resetRevision, lifetime]);
 *   const selected = providers.find(provider => provider.label === selection) ?? providers.find(provider => provider.label === 'SPFxApplicationCustomizerProvider') ?? providers[0];
 *   const available = typeof selected.value === 'function';
 *   const regions: DemoRegions = {
 *     controls: <Field label="Provider"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'providers' } }} value={selection} onChange={(_, data) => {
 *       if (lifetime.active && providers.some(provider => provider.label === data.value)) setSelection(data.value);
 *     }}>{providers.map(provider => <option key={provider.label} value={provider.label}>{provider.label}</option>)}</Select></Field>,
 *     children: <div className={styles.result} data-providers-demo="result">
 *       <Badge appearance="tint" color={available ? 'success' : 'danger'}>{selected.label} export</Badge>
 *       <dl className={styles.facts}>
 *         <dt>Export available</dt><dd>{available ? 'Yes' : 'No'}</dd>
 *         <dt>Current host</dt><dd>{context.kind}</dd>
 *         <dt>Required host</dt><dd>{selected.kind}</dd>
 *         <dt>Host match</dt><dd>{context.kind === selected.kind ? 'Yes' : 'No: requires another host'}</dd>
 *       </dl>
 *     </div>,
 *     diagnostics: <p>Each provider requires its own initialized SPFx host. Selection does not mount new instances or query services. Export availability and host match do not establish native execution. Field Customizer and Command Set native lifecycle: NOT EXECUTED by this diagnostic.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}<Accordion collapsible><AccordionItem value="diagnostics"><AccordionHeader>Advanced diagnostics</AccordionHeader><AccordionPanel>{regions.diagnostics}</AccordionPanel></AccordionItem></Accordion></>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Choose each of four providers and compare the actual exported function, required kind and current host independently. Reset restores this owner’s provider choice without mounting a new provider.' };
 * ```

 * @scenario Inspect the actual Top placeholder provider
 * ```tsx
 * import * as React from 'react';
 * import { Accordion, AccordionHeader, AccordionItem, AccordionPanel, MessageBar, MessageBarBody } from '@fluentui/react-components';
 * import { useSPFxInstanceInfo, useSPFxPageContext, useSPFxProperties } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type DemoProps = { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number; applicationCustomizerTop?: { readonly placeholder: 'Top' } };
 * function present(regions: DemoRegions, render?: DemoProps['render']): React.ReactElement {
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}<Accordion collapsible><AccordionItem value="diagnostics"><AccordionHeader>Advanced diagnostics</AccordionHeader><AccordionPanel>{regions.diagnostics}</AccordionPanel></AccordionItem></Accordion></>;
 * }
 * export function Demo({ render, applicationCustomizerTop }: DemoProps = {}): React.ReactElement {
 *   if (!applicationCustomizerTop) return present({ status: <MessageBar><MessageBarBody>This observation requires the actual initialized Application Customizer in its native Top placeholder.</MessageBarBody></MessageBar> }, render);
 *   return <ProviderObservation render={render} />;
 * }
 * function ProviderObservation({ render }: Pick<DemoProps, 'render'>): React.ReactElement {
 *   const { id, kind } = useSPFxInstanceInfo();
 *   if (kind !== 'AppCustomizer') return present({ status: <MessageBar intent="warning"><MessageBarBody>Top observation requires AppCustomizer; the actual provider is {kind}.</MessageBarBody></MessageBar> }, render);
 *   return <TopObservation id={id} kind={kind} render={render} />;
 * }
 * function TopObservation({ id, kind, render }: { id: string; kind: string; render?: DemoProps['render'] }): React.ReactElement {
 *   const pageContext = useSPFxPageContext();
 *   const { properties } = useSPFxProperties<{ testMessage: string }>();
 *   return present({
 *     children: <p data-providers-demo="native-top">{`SPFxApplicationCustomizerProvider active: ${kind} ${id} | ${pageContext.web.title} | ${properties?.testMessage ?? 'No message'}`}</p>,
 *     diagnostics: <p>This read-only observation uses the matching initialized provider owned by the native Top root. Reset has no local state to clear and never writes properties, replaces a provider or disposes the placeholder. Properties follow the provider’s shallow host-render reconciliation.</p>,
 *   }, render);
 * }
 * export const demo = { host: 'applicationCustomizer', hostAdapter: 'applicationCustomizerTop', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual AppCustomizer kind, instance ID, web title and testMessage in Top. An absent message shows No message; an empty string remains empty. Reset does not write the native host.' };
 * ```
 */
export function SPFxApplicationCustomizerProvider<TProps extends {} = {}>(
  props: SPFxApplicationCustomizerProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
