/**
 * Mount a native ListView Command Set React surface with an isolated toolkit runtime.
 *
 * @packageDocumentation
 * @docPage core/providers
 */
// provider-listview-commandset.tsx
// Type-safe provider for ListView Command Sets

import * as React from 'react';
import type { BaseListViewCommandSet } from '@microsoft/sp-listview-extensibility';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Initialized native host instance and optional subtree for SPFxListViewCommandSetProvider. Each mounted provider owns a separate runtime store. Children mount after the active ServiceScope finishes; missing-scope fallback does not make native services available.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @example Typed native host input
 * ```tsx
 * import * as React from 'react';
 * import { BaseListViewCommandSet } from '@microsoft/sp-listview-extensibility';
 * import { SPFxListViewCommandSetProvider, SPFxListViewCommandSetProviderProps } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(instance: BaseListViewCommandSet<{ title: string }>): React.ReactElement {
 *   const props: SPFxListViewCommandSetProviderProps<{ title: string }> = { instance, children: <span>Host content</span> };
 *   return <SPFxListViewCommandSetProvider {...props} />;
 * }
 * ```
 */
export interface SPFxListViewCommandSetProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized ListView Command Set instance supplying context, properties and lifecycle integration.
   */
  instance: BaseListViewCommandSet<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omitted or undefined children render no application content.
   */
  children?: React.ReactNode;
}

/**
 * Establishes SPFx context and isolated runtime state for a ListView Command Set React surface. Pass its initialized native instance. Descendants mount after the active ServiceScope finishes; the provider never registers, finishes or authorizes services. The missing-scope fallback allows rendering but service-consuming hooks can fail.
 *
 * Host renders reconcile shallow property snapshots, including removals and reference changes; nested in-place changes are not deep-observed. Hook property updates synchronize back to the native instance. Theme subscriptions and runtime listeners are cleaned up on unmount; obsolete scope-readiness callbacks are ignored. Keep a mounted provider associated with its owning host surface and unmount React when that surface is disposed.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @param props - Required initialized BaseListViewCommandSet instance and optional children; omitted or undefined children render no application content.
 * @returns Provider element establishing context after ServiceScope readiness.
 * @useWhen toolkit hooks render inside a native ListViewCommandSet surface owned by this host instance.
 * @avoidWhen the host is another SPFx component kind; use its matching typed provider.
 * @example Mount an initialized native host
 * ```tsx
 * import * as React from 'react';
 * import { BaseListViewCommandSet } from '@microsoft/sp-listview-extensibility';
 * import {
 *   SPFxListViewCommandSetProvider,
 *   SPFxListViewCommandSetProviderProps
 * } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(
 *   instance: BaseListViewCommandSet<{ title: string }>
 * ): React.ReactElement {
 *   const props: SPFxListViewCommandSetProviderProps<{ title: string }> = {
 *     instance,
 *     children: <span>Host content</span>
 *   };
 *   return <SPFxListViewCommandSetProvider {...props} />;
 * }
 * ```
 * @public
 * @docPage core/providers
 *
 *
 * @example Dispose a temporary command dialog root
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { BaseListViewCommandSet, type IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
 * import { SPFxListViewCommandSetProvider } from '@apvee/spfx-react-toolkit';
 * function DetailsDialog({ count, onClose }: { count: number; onClose: () => void }) {
 *   return <section role="dialog" aria-label="Selected rows"><p>{count} selected</p><button onClick={onClose}>Close</button></section>;
 * }
 * export default class DetailsCommand extends BaseListViewCommandSet<{ dialogTitle: string }> {
 *   private closeDialog?: () => void;
 *   public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
 *     if (event.itemId !== 'SHOW_DETAILS') return;
 *     this.closeDialog?.();
 *     const root = document.createElement('div');
 *     document.body.appendChild(root);
 *     const close = () => {
 *       ReactDom.unmountComponentAtNode(root);
 *       root.remove();
 *       if (this.closeDialog === close) this.closeDialog = undefined;
 *     };
 *     this.closeDialog = close;
 *     ReactDom.render(<SPFxListViewCommandSetProvider instance={this}>
 *       <DetailsDialog count={this.context.listView.selectedRows?.length ?? 0} onClose={close} />
 *     </SPFxListViewCommandSetProvider>, root);
 *   }
 *   public onDispose(): void { this.closeDialog?.(); }
 * }
 * ```

 * @scenario Compare Command Set provider availability with the current host
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
 *   const [selection, setSelection] = React.useState('SPFxListViewCommandSetProvider');
 *   const revision = React.useRef(resetRevision);
 *   const lifetime = React.useMemo(() => ({ active: true }), [resetRevision]);
 *   React.useLayoutEffect(() => {
 *     if (revision.current !== resetRevision) {
 *       revision.current = resetRevision;
 *       setSelection('SPFxListViewCommandSetProvider');
 *     }
 *     return () => { lifetime.active = false; };
 *   }, [resetRevision, lifetime]);
 *   const selected = providers.find(provider => provider.label === selection) ?? providers.find(provider => provider.label === 'SPFxListViewCommandSetProvider') ?? providers[0];
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
 */
export function SPFxListViewCommandSetProvider<TProps extends {} = {}>(
  props: SPFxListViewCommandSetProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
