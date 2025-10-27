// useSPFxPerformance.ts
// Hook for performance measurement and monitoring

import { useCallback } from 'react';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';
import { useSPFxCorrelationInfo } from './useSPFxCorrelationInfo';

/**
 * Performance result type
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
 */
export interface SPFxPerformanceInfo {
  /** Create a performance mark */
  readonly mark: (name: string) => void;
  
  /** Measure duration between two marks */
  readonly measure: (name: string, startMark: string, endMark?: string) => SPFxPerfResult;
  
  /** Time an async operation */
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
 *   const [data, setData] = useState(null);
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
      const startMark = name + '-start';
      mark(startMark);
      
      const result = await Promise.resolve().then(fn);
      
      const measurement = measure(name, startMark);
      
      return {
        ...measurement,
        result,
      };
    },
    [mark, measure]
  );
  
  return {
    mark,
    measure,
    time,
  };
}
