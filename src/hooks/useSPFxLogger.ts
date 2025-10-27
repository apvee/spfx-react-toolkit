// useSPFxLogger.ts
// Hook for structured logging with SPFx context

import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';
import { useSPFxUserInfo } from './useSPFxUserInfo';
import { useSPFxSiteInfo } from './useSPFxSiteInfo';
import { useSPFxCorrelationInfo } from './useSPFxCorrelationInfo';

/**
 * Log levels
 */
export type LogLevel = 'debug' | 'info' | 'warn' | 'error';

/**
 * Structured log entry
 */
export interface LogEntry {
  /** Log level */
  readonly level: LogLevel;
  
  /** Log message */
  readonly message: string;
  
  /** Timestamp ISO string */
  readonly ts: string;
  
  /** SPFx instance ID */
  readonly instanceId: string;
  
  /** Host kind */
  readonly host: string;
  
  /** Current user */
  readonly user: string;
  
  /** Site collection URL */
  readonly siteUrl: string | undefined;
  
  /** Web URL */
  readonly webUrl: string | undefined;
  
  /** Correlation ID */
  readonly correlationId: string | undefined;
  
  /** Extra metadata */
  readonly extra?: Record<string, unknown>;
}

/**
 * Return type for useSPFxLogger hook
 */
export interface SPFxLoggerInfo {
  /** Log debug message */
  readonly debug: (message: string, extra?: Record<string, unknown>) => void;
  
  /** Log info message */
  readonly info: (message: string, extra?: Record<string, unknown>) => void;
  
  /** Log warning message */
  readonly warn: (message: string, extra?: Record<string, unknown>) => void;
  
  /** Log error message */
  readonly error: (message: string, extra?: Record<string, unknown>) => void;
}

/**
 * Hook for structured logging with SPFx context
 * 
 * Provides structured logging methods that automatically include:
 * - SPFx instance ID
 * - Host kind (WebPart, Extension, etc.)
 * - Current user information
 * - Site/web URLs
 * - Correlation ID
 * - Timestamp
 * 
 * By default logs to console, but can be configured with custom handler
 * for integration with Application Insights, Log Analytics, or other
 * logging services.
 * 
 * Useful for:
 * - Diagnostic logging
 * - Error tracking
 * - Performance monitoring
 * - User activity tracking
 * - Support troubleshooting
 * 
 * @param handler - Optional custom log handler
 * @returns Logger methods
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const logger = useSPFxLogger();
 *   
 *   const handleClick = () => {
 *     logger.info('Button clicked', { buttonId: 'save' });
 *   };
 *   
 *   const handleError = (error: Error) => {
 *     logger.error('Operation failed', {
 *       errorMessage: error.message,
 *       stack: error.stack
 *     });
 *   };
 *   
 *   return <button onClick={handleClick}>Click Me</button>;
 * }
 * ```
 * 
 * @example Custom handler for Application Insights
 * ```tsx
 * const customHandler = (entry: LogEntry) => {
 *   appInsights.trackTrace({
 *     message: entry.message,
 *     severityLevel: entry.level,
 *     properties: {
 *       instanceId: entry.instanceId,
 *       correlationId: entry.correlationId,
 *       ...entry.extra
 *     }
 *   });
 * };
 * 
 * const logger = useSPFxLogger(customHandler);
 * ```
 */
export function useSPFxLogger(
  handler?: (entry: LogEntry) => void
): SPFxLoggerInfo {
  const { id: instanceId, kind: host } = useSPFxInstanceInfo();
  const { displayName, loginName } = useSPFxUserInfo();
  const { siteUrl, webUrl } = useSPFxSiteInfo();
  const { correlationId } = useSPFxCorrelationInfo();
  
  const emit = (level: LogLevel, message: string, extra?: Record<string, unknown>): void => {
    const entry: LogEntry = {
      level,
      message,
      ts: new Date().toISOString(),
      instanceId,
      host,
      user: displayName + ' (' + loginName + ')',
      siteUrl,
      webUrl,
      correlationId,
      extra,
    };
    
    if (handler) {
      handler(entry);
    } else {
      // Default: log to console
      const levelUpper = level.toUpperCase();
      const line = '[' + levelUpper + '] ' + entry.ts + ' ' + host + '/' + instanceId + ' – ' + message;
      
      // Use appropriate console method
      const consoleFn = level === 'debug' ? console.log : console[level];
      consoleFn(line, extra ?? {});
    }
  };
  
  return {
    debug: (m: string, e?: Record<string, unknown>): void => emit('debug', m, e),
    info: (m: string, e?: Record<string, unknown>): void => emit('info', m, e),
    warn: (m: string, e?: Record<string, unknown>): void => emit('warn', m, e),
    error: (m: string, e?: Record<string, unknown>): void => emit('error', m, e),
  };
}
