/**
 * Browser measurements carry provider identity and optional correlation metadata.
 *
 * @packageDocumentation
 * @docPage hooks/performance
 */
// useSPFxPerformance.ts
// Hook for performance measurement and monitoring

import { useCallback, useMemo } from 'react';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';
import { useSPFxCorrelationInfo } from './useSPFxCorrelationInfo';

// A scalar sequence avoids collisions across hook instances without retaining operations.
let timedOperationSequence = 0;

/**
 * Named browser measurement with duration/context and an optional timed callback result.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPerfResult, useSPFxPerformance } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const perf = useSPFxPerformance();
 *   const [pending, setPending] = React.useState(false);
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       setPending(true);
 *       const timed: SPFxPerfResult<Response> = await perf.time('load', () => fetch('/api/data'));
 *       setResult({ durationMs: timed.durationMs, status: timed.result?.status });
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     } finally {
 *       setPending(false);
 *     }
 *   };
 *   return <>
 *     <button disabled={pending} onClick={() => void run()}>Load data</button>
 *     {failure && <span role="alert">{failure}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 * @typeParam T - Actual value returned by the timed callback, defaulting to unknown.
 */
export interface SPFxPerfResult<T = unknown> {
  /** Measurement name */
  readonly name: string;
  
  /** Duration in milliseconds */
  readonly durationMs: number;
  
  /** Callback result from time; omitted from plain measure results, and may itself be undefined. */
  readonly result?: T;
  
  /** SPFx instance ID */
  readonly instanceId: string;
  
  /** Host kind */
  readonly host: string;
  
  /** Correlation ID for tracking */
  readonly correlationId: string | undefined;
}

/**
 * Browser measurement methods enriched with instance, host and optional correlation metadata.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPerformanceInfo, useSPFxPerformance } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const perf: SPFxPerformanceInfo = useSPFxPerformance();
 *   const [pending, setPending] = React.useState(false);
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       setPending(true);
 *       const timed = await perf.time('load', () => fetch('/api/data'));
 *       setResult({ durationMs: timed.durationMs, status: timed.result?.status });
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     } finally {
 *       setPending(false);
 *     }
 *   };
 *   return <>
 *     <button disabled={pending} onClick={() => void run()}>Load data</button>
 *     {failure && <span role="alert">{failure}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 */
export interface SPFxPerformanceInfo {
  /**
   * Creates a global browser performance mark when supported; mark errors are ignored. Callers own cleanup of explicit marks.
   * @param name - Global mark name.
   * @returns No value.
   */
  readonly mark: (name: string) => void;
  
  /**
   * Measures from startMark to a newly created endMark, or now when omitted. Does not clear explicit marks or measurement entries.
   * @param name - Measurement name.
   * @param startMark - Existing start mark.
   * @param endMark - Name for a new end mark created at measurement time; omitted or undefined measures to the current time.
   * @returns Context-enriched measurement; durationMs is zero if unsupported or measurement fails.
   */
  readonly measure: (name: string, startMark: string, endMark?: string) => SPFxPerfResult;
  
  /**
   * Times a synchronous or asynchronous callback and clears its unique start mark after settlement. Measurement entries remain in the browser timeline.
   * @param name - Measurement name.
   * @typeParam T - Callback result type.
   * @param fn - Operation invoked asynchronously; errors reject the returned promise.
   * @returns Duration, context and callback result.
   * @throws Callback errors.
   */
  readonly time: <T>(name: string, fn: /**
   * Timed work, invoked in a promise continuation.
   * @returns Synchronous value or promise for the timed result.
   */ () => Promise<T> | T) => Promise<SPFxPerfResult<T>>;
}

/**
 * Hook for performance measurement and monitoring
 *
 * Provides access to the Performance API for measuring code execution time.
 *
 * Methods:
 * - mark(): Create named performance marks
 * - measure(): Calculate duration between marks
 * - time(): Wrap async operations with automatic timing
 *
 * All measurements include SPFx context (instanceId, correlationId) for
 * integration with logging and monitoring systems.
 *
 * Useful for:
 * - Performance profiling
 * - Identifying bottlenecks
 * - Monitoring real-world performance
 * - Integration with Application Insights
 * - Custom performance dashboards
 *
 * @returns Performance measurement methods
 *
 * @example Measure a requested data load
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPerformance } from '@apvee/spfx-react-toolkit';
 *
 * function MyComponent() {
 *   const { time } = useSPFxPerformance();
 *   const [pending, setPending] = React.useState(false);
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       setPending(true);
 *       const timed = await time('fetch-data', async () => {
 *         const response = await fetch('/api/data');
 *         return response.json();
 *       });
 *       setResult({ durationMs: timed.durationMs, data: timed.result });
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     } finally {
 *       setPending(false);
 *     }
 *   };
 *   return (
 *     <>
 *       <button disabled={pending} onClick={() => void run()}>
 *         Load data
 *       </button>
 *       {failure && <span role="alert">{failure}</span>}
 *       <pre>{JSON.stringify(result)}</pre>
 *     </>
 *   );
 * }
 * ```
 * @throws When called outside an SPFx provider.
 *
 * @useWhen diagnostics need timed operations enriched with host and correlation metadata.
 *
 * @scenario Measure a cancellable local timed operation
 * ```tsx
 * // Requires the matching SPFx provider. Only Run starts timing. Local cancellation settles the owned await and suppresses stale success or failure, including cancellation before the timed callback begins. The hook clears its own unique start mark; measurement entries remain. No global timeline cleanup occurs.
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import { useSPFxContext, useSPFxPerformance } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   children?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * interface Run {
 *   timer?: number;
 *   settle?: () => void;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const { spfxContext } = useSPFxContext();
 *   const performance = useSPFxPerformance();
 *   const [pending, setPending] = React.useState(false);
 *   const [result, setResult] = React.useState<string>();
 *   const [duration, setDuration] = React.useState<number>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const lifetime = React.useMemo(() => ({ active: true, run: undefined as Run | undefined }),
 *     [resetRevision, spfxContext, performance]);
 *   React.useLayoutEffect(() => {
 *     setPending(false);
 *     setResult(undefined);
 *     setDuration(undefined);
 *     setFailure(undefined);
 *     return () => {
 *       lifetime.active = false;
 *       const run = lifetime.run;
 *       lifetime.run = undefined;
 *       if (run?.timer !== undefined) window.clearTimeout(run.timer);
 *       // Cancelling our delay must also release the promise owned by time().
 *       run?.settle?.();
 *     };
 *   }, [lifetime]);
 *   const runOperation = async (): Promise<void> => {
 *     if (!lifetime.active || lifetime.run) return;
 *     const run: Run = {};
 *     lifetime.run = run;
 *     const current = (): boolean => lifetime.active && lifetime.run === run;
 *     const finish = (): void => {
 *       if (current()) { lifetime.run = undefined; setPending(false); }
 *     };
 *     setPending(true);
 *     setFailure(undefined);
 *     setResult(undefined);
 *     setDuration(undefined);
 *     try {
 *       const timed = await performance.time('spfx-toolkit-demo', async () => {
 *         // time invokes this in a continuation; Reset may already have cancelled it.
 *         if (!current()) return undefined;
 *         await new Promise<void>(resolve => {
 *           run.settle = resolve;
 *           run.timer = window.setTimeout(() => {
 *             run.timer = undefined;
 *             run.settle = undefined;
 *             resolve();
 *           }, 250);
 *         });
 *         return current() ? 'Timed operation completed' : undefined;
 *       });
 *       if (current()) {
 *         setResult(timed.result);
 *         setDuration(timed.durationMs);
 *       }
 *     } catch (error) {
 *       if (current()) setFailure(error instanceof Error ? error.message : String(error));
 *     } finally {
 *       finish();
 *     }
 *   };
 *   return render({
 *     commands: <><Button disabled={pending} onClick={runOperation}>Run timed operation</Button></>,
 *     status: failure ? <p role="alert">{failure}</p> : pending ? <p role="status">Timing operation…</p> : undefined,
 *     children: <output>Last Result: {result ?? ''}</output>,
 *     diagnostics: duration !== undefined ? <p>Duration: {duration.toFixed(2)}ms</p> : undefined,
 *   });
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run measures a real local 250 ms wait and shows the actual result and duration (which may be zero when unsupported). Reset, host replacement and departure cancel the owned wait and suppress obsolete results; they do not clear unrelated browser timeline entries.' };
 * ```
 */
export function useSPFxPerformance(): SPFxPerformanceInfo {
  const { id: instanceId, kind: host } = useSPFxInstanceInfo();
  const { correlationId } = useSPFxCorrelationInfo();
  
  // Create a performance mark
  const mark = useCallback((name: string): void => {
    if (typeof performance?.mark === 'function') {
      try {
        performance.mark(name);
      } catch {
        // Swallow errors (e.g., invalid mark name)
      }
    }
  }, []);
  
  // Measure duration between marks
  const measure = useCallback(
    (name: string, startMark: string, endMark?: string): SPFxPerfResult => {
      try {
        // Create end mark if provided
        if (endMark && typeof performance?.mark === 'function') {
          performance.mark(endMark);
        }
        
        // Create measure
        if (typeof performance?.measure === 'function') {
          performance.measure(name, startMark, endMark);
          
          // Get the measurement
          const entries = performance.getEntriesByName(name, 'measure');
          const lastEntry = entries[entries.length - 1];
          
          return {
            name,
            durationMs: lastEntry?.duration ?? 0,
            instanceId,
            host,
            correlationId,
          };
        }
      } catch {
        // Swallow errors
      }
      
      // Fallback result
      return {
        name,
        durationMs: 0,
        instanceId,
        host,
        correlationId,
      };
    },
    [instanceId, host, correlationId]
  );
  
  // Time an async operation
  const time = useCallback(
    async <T,>(name: string, fn: /**
   * Timed work, invoked in a promise continuation.
   * @returns Synchronous value or promise for the timed result.
   */ () => Promise<T> | T): Promise<SPFxPerfResult<T>> => {
      const startMark = name + '-start-' + (++timedOperationSequence);
      mark(startMark);
      
      try {
        const result = await Promise.resolve().then(fn);
        const measurement = measure(name, startMark);
        return {
          ...measurement,
          result,
        };
      } finally {
        if (typeof performance?.clearMarks === 'function') {
          performance.clearMarks(startMark);
        }
      }
    },
    [mark, measure]
  );
  
  return useMemo(() => ({
    mark,
    measure,
    time,
  }), [mark, measure, time]);
}
