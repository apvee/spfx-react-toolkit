// useSPFxPerformance.ts
// Hook for performance measurement and monitoring

import { useCallback, useMemo } from 'react';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';
import { useSPFxCorrelationInfo } from './useSPFxCorrelationInfo';

// A scalar sequence avoids collisions across hook instances without retaining operations.
let timedOperationSequence = 0;

/**
 * Performance result type
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const perf = useSPFxPerformance();
 * const onLoad = async () => {
 *   const result: SPFxPerfResult<Response> = await perf.time('load', () => fetch('/api/data'));
 *   console.log(result.durationMs, result.result?.status);
 * };
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxPerfResult<T = unknown> {
  /** Measurement name */
  readonly name: string;
  
  /** Duration in milliseconds */
  readonly durationMs: number;
  
  /** Optional result from timed operation */
  readonly result?: T;
  
  /** SPFx instance ID */
  readonly instanceId: string;
  
  /** Host kind */
  readonly host: string;
  
  /** Correlation ID for tracking */
  readonly correlationId: string | undefined;
}

/**
 * Return type for useSPFxPerformance hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const perf: SPFxPerformanceInfo = useSPFxPerformance();
 * const onLoad = async () => console.log(await perf.time('load', () => fetch('/api/data')));
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
   * @param endMark - Optional name for a new end mark created at measurement time.
   * @returns Context-enriched measurement; durationMs is zero if unsupported or measurement fails.
   */
  readonly measure: (name: string, startMark: string, endMark?: string) => SPFxPerfResult;
  
  /**
   * Times a synchronous or asynchronous callback and clears its unique start mark after settlement. Measurement entries remain in the browser timeline.
   * @param name - Measurement name.
   * @param fn - Operation invoked asynchronously; errors reject the returned promise.
   * @returns Duration, context and callback result.
   * @throws Callback errors.
   */
  readonly time: <T>(name: string, fn: () => Promise<T> | T) => Promise<SPFxPerfResult<T>>;
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
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { mark, measure, time } = useSPFxPerformance();
 *   const [data, setData] = useState<unknown>(undefined);
 *   
 *   const fetchData = async () => {
 *     const result = await time('fetch-data', async () => {
 *       const response = await fetch('/api/data');
 *       return response.json();
 *     });
 *     
 *     console.log(`Fetch took ${result.durationMs}ms`);
 *     setData(result.result);
 *   };
 *   
 *   return <button onClick={fetchData}>Load Data</button>;
 * }
 * ```
 * @throws When called outside an SPFx provider.
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
    async <T,>(name: string, fn: () => Promise<T> | T): Promise<SPFxPerfResult<T>> => {
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
