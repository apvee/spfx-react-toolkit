/**
 * Invoke caller work and execute PnP batches on an explicitly configured client.
 *
 * @packageDocumentation
 * @docPage services/pnp
 */
import type { SPFI } from '@pnp/sp';

import '@pnp/sp/batching';

/**
 * Explicit operations on a caller-owned configured SPFI; arbitrary callbacks require their own PnP feature registrations. Registers batching only and adds no caching, credentials or React state.
 *
 * @example
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import '@pnp/sp/webs';
 * import { SPFxPnPService, createSPFxPnPService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const service: SPFxPnPService = createSPFxPnPService(sp);
 *   const web = await service.invoke(client => client.web());
 * }
 * ```
 */
export interface SPFxPnPService {
  /**
   * The exact caller-supplied PnP client.
   */
  readonly sp: SPFI;
  /**
   * Runs a callback with the supplied client.
   * @typeParam T - Callback result, inferred from the returned Promise.
   * @param fn - Required callback returning a Promise; the service observes its result and rejection.
   * @returns The callback result.
   * @throws Callback or request failures.
   */
  invoke: <T>(fn: /**
     * Runs caller work on the original client.
     * @param sp - Exact caller-supplied configured SPFI.
     * @returns Promise for the caller's operation result.
     */ (sp: SPFI) => Promise<T>) => Promise<T>;
  /**
   * Invokes the callback to queue work, executes the batch, then awaits its result. Queue operations before awaiting their results in the callback: requests queued after an await can miss execution. Batches are not transactions and successful writes are not rolled back.
   * @typeParam T - Callback result, inferred from the returned Promise.
   * @param fn - Required callback returning a Promise; the service observes its result and rejection.
   * @returns The callback result.
   * @throws Callback or request failures; batch execution failures also reject.
   * @example
   * ```ts
   * import type { SPFxPnPService } from '@apvee/spfx-react-toolkit';
   * import '@pnp/sp/webs';
   * async function readTogether(service: SPFxPnPService) {
   *   return service.batch(client => {
   *     const first = client.web.select('Title')();
   *     const second = client.web.select('Url')();
   *     return Promise.all([first, second]);
   *   });
   * }
   * ```
   */
  batch: <T>(fn: /**
     * Queues operations before the service executes the batch.
     * @param batchedSP - PnP batch client derived from the original SPFI.
     * @returns Promise for queued results; queue requests synchronously before awaiting them.
     */ (batchedSP: SPFI) => Promise<T>) => Promise<T>;
}

/**
 * Creates generic operations without React state; no request is made at construction.
 * @param sp - Authenticated caller-owned PnP client.
 * @returns The original client and explicit invoke/batch methods.
 *
 * @example Basic usage
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import '@pnp/sp/webs';
 * import { createSPFxPnPService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const service = createSPFxPnPService(sp);
 *   const web = await service.invoke(client => client.web());
 * }
 * ```
 * @useWhen You already own an authenticated SPFI and need explicit arbitrary invoke or batch callbacks outside React.
 *
 *
 * @example Compare invoke batch and client identity
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import '@pnp/sp/webs';
 * import { createSPFxPnPService } from '@apvee/spfx-react-toolkit/services';
 * async function demonstrate(sp: SPFI) {
 *
 *   const service = createSPFxPnPService(sp); // Authenticated SPFI supplied by the caller.
 *   const currentWeb = await service.invoke(client => client.web());
 *   const batchedWeb = await service.batch(client => client.web());
 *   console.log(service.sp === sp, currentWeb, batchedWeb); // true
 * }
 * ```
 *
 * @scenario Read through a supplied PnP service client
 * ```tsx
 * import * as React from 'react';
 * import { Button, makeStyles } from '@fluentui/react-components';
 * import { useSPFxContext, useSPFxPageContext, createSPFxPnPContextService, createSPFxPnPService } from '@apvee/spfx-react-toolkit';
 * import '@pnp/sp/webs';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ controls: { display: 'flex', flexDirection: 'column', gap: '12px' } });
 * type Data = { kind: 'invoke' | 'batch'; title: string; url: string; description?: string };
 * type Outcome = { token: object; request: number } & (
 *   | { phase: 'pending' }
 *   | { phase: 'success'; data: Data }
 *   | { phase: 'error'; message: string }
 * );
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const pageContext = useSPFxPageContext();
 *   const webUrl = pageContext.web.absoluteUrl;
 *   const configured = React.useMemo(() => {
 *     try {
 *       const sp = createSPFxPnPContextService(spfxContext, pageContext).createSPFI();
 *       return { sp, service: createSPFxPnPService(sp), error: undefined };
 *     } catch (error) {
 *       return { sp: undefined, service: undefined, error: error instanceof Error ? error.message : String(error) };
 *     }
 *   }, [spfxContext, pageContext, webUrl]);
 *   const values: readonly unknown[] = [configured, resetRevision];
 *   const boundary = React.useRef({ values, token: {} });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(false);
 *   const sequence = React.useRef(0);
 *   const dispatch = React.useRef<{ token: object; request: number }>();
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const current = (): boolean => mounted.current && boundary.current.token === token;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; sequence.current++; }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     sequence.current++;
 *     setOutcome(undefined);
 *   }, [resetRevision]);
 *   const run = async (kind: 'invoke' | 'batch'): Promise<void> => {
 *     const service = configured.service;
 *     if (!current() || !service || dispatch.current?.token === token) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     dispatch.current = { token, request };
 *     setOutcome({ token, request, phase: 'pending' });
 *     try {
 *       let data: Data;
 *       if (kind === 'invoke') {
 *         // SharePoint's raw legacy text field may be null; normalize local absence.
 *         // eslint-disable-next-line @rushstack/no-new-null
 *         const web: { Title: string; Url: string; Description?: string | null } = await service.invoke(client => client.web.select('Title', 'Url', 'Description')());
 *         data = { kind, title: web.Title, url: web.Url, description: web.Description ?? undefined };
 *       } else {
 *         const [title, url]: [{ Title: string }, { Url: string }] = await service.batch(client => {
 *           // Queue BOTH operations before returning/awaiting either result.
 *           const first = client.web.select('Title')();
 *           const second = client.web.select('Url')();
 *           return Promise.all([first, second]);
 *         });
 *         data = { kind, title: title.Title, url: url.Url };
 *       }
 *       if (owns()) setOutcome({ token, request, phase: 'success', data });
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (owns() && dispatch.current?.request === request && dispatch.current.token === token) dispatch.current = undefined;
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (!current() || visible?.phase !== 'error' || visible.request !== sequence.current) return;
 *     setOutcome(undefined);
 *   };
 *   const regions: DemoRegions = {
 *     controls: <div className={classes.controls}><p>Configured web: {webUrl}</p><p>Exact supplied client retained: {configured.service ? String(configured.service.sp === configured.sp) : 'Unavailable'}</p></div>,
 *     commands: <><Button disabled={!configured.service || visible?.phase === 'pending'} onClick={() => run('invoke')}>Invoke web read</Button><Button disabled={!configured.service || visible?.phase === 'pending'} onClick={() => run('batch')}>Batch web reads</Button></>,
 *     status: <>
 *       {configured.error && <p role="alert">Client initialization error: {configured.error}</p>}
 *       {visible?.phase === 'error' && <div role="alert">Request error: {visible.message} <Button onClick={dismiss}>Dismiss</Button></div>}
 *       <p role="status">{visible?.phase === 'pending' ? 'Reading through service...' : visible?.phase === 'success' ? 'Service read complete.' : 'Choose an explicit service command.'}</p>
 *     </>,
 *     children: visible?.phase === 'success' ? <dl aria-label="Supplied client service result"><dt>Command</dt><dd>{visible.data.kind}</dd><dt>Title</dt><dd>{visible.data.title}</dd><dt>URL</dt><dd>{visible.data.url}</dd>{visible.data.description && <><dt>Description</dt><dd>{visible.data.description}</dd></>}</dl> : undefined,
 *     diagnostics: <p>Mount below the matching SPFx provider. The caller configures authentication; the service retains that exact SPFI and adds no React loading/error state. Each command is explicit. Batch queues Title then Url before awaiting, and is not a transaction. This Demo owns pending/error/currentness and consumes action failures. Reset/client/departure retire old commands and results without canceling dispatched work. No writes or automatic chained commands run.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Compare explicit invoke and queued batch web reads while the service retains the exact supplied client and obsolete commands cannot publish results.' };
 * ```
 */
export function createSPFxPnPService(sp: SPFI): SPFxPnPService {
  const invoke = async <T>(fn: /**
     * Runs caller work on the original client.
     * @param sp - Exact caller-supplied configured SPFI.
     * @returns Promise for the caller's operation result.
     */ (sp: SPFI) => Promise<T>): Promise<T> => {
    return fn(sp);
  };

  const batch = async <T>(fn: /**
     * Queues operations before the service executes the batch.
     * @param batchedSP - PnP batch client derived from the original SPFI.
     * @returns Promise for queued results; queue requests synchronously before awaiting them.
     */ (batchedSP: SPFI) => Promise<T>): Promise<T> => {
    const [batchedSP, execute] = sp.batched();
    const resultPromise = fn(batchedSP);

    await execute();

    return resultPromise;
  };

  return {
    sp,
    invoke,
    batch
  };
}
