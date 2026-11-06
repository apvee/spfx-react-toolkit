/**
 * Synchronous log sinks receive structured host and user diagnostics.
 *
 * @packageDocumentation
 * @docPage hooks/performance
 */
// useSPFxLogger.ts
// Hook for structured logging with SPFx context

import { useCallback, useMemo } from 'react';
import { useSPFxContext } from './useSPFxContext';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';
import { useSPFxUserInfo } from './useSPFxUserInfo';
import { useSPFxSiteInfo } from './useSPFxSiteInfo';
import { useSPFxCorrelationInfo } from './useSPFxCorrelationInfo';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Log levels
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { LogLevel } from '@apvee/spfx-react-toolkit';
 *
 * const level: LogLevel = 'warn';
 * ```
 */
export type LogLevel = 'debug' | 'info' | 'warn' | 'error';

/**
 * Timestamped message and host/user/site/correlation metadata emitted synchronously to a custom logger sink.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { LogEntry, useSPFxLogger } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const handleLog = (entry: LogEntry) => console.info(entry.message, entry.instanceId);
 *   const logger = useSPFxLogger(handleLog);
 *   return <button onClick={() => logger.info('Selected')}>Log</button>;
 * }
 * ```
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
  
  /** WebPart tag; undefined for other hosts or when missing from context */
  readonly webPartTag?: string;
  
  /** Extra metadata; undefined when omitted by the caller. */
  readonly extra?: Record<string, unknown>;
}

/**
 * Synchronous structured log methods using the configured sink or console, enriched with provider metadata.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxLoggerInfo, useSPFxLogger } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const logger: SPFxLoggerInfo = useSPFxLogger();
 *   const onSave = () => logger.info('Saved', { itemId: 12 });
 *   return <button onClick={()=>void onSave()}>Run operation</button>;
 * }
 * ```
 */
export interface SPFxLoggerInfo {
  /**
   * Emits a debug message to the custom handler, or the corresponding console output (debug uses console.log). Handler failures propagate.
   * @param message - Log message.
   * @param extra - Metadata included in the entry; omitted or undefined adds no entry metadata.
   * @returns No value.
   */
  readonly debug: (message: string, extra?: Record<string, unknown>) => void;
  
  /**
   * Emits a info message to the custom handler, or the corresponding console output (debug uses console.log). Handler failures propagate.
   * @param message - Log message.
   * @param extra - Metadata included in the entry; omitted or undefined adds no entry metadata.
   * @returns No value.
   */
  readonly info: (message: string, extra?: Record<string, unknown>) => void;
  
  /**
   * Emits a warn message to the custom handler, or the corresponding console output (debug uses console.log). Handler failures propagate.
   * @param message - Log message.
   * @param extra - Metadata included in the entry; omitted or undefined adds no entry metadata.
   * @returns No value.
   */
  readonly warn: (message: string, extra?: Record<string, unknown>) => void;
  
  /**
   * Emits a error message to the custom handler, or the corresponding console output (debug uses console.log). Handler failures propagate.
   * @param message - Log message.
   * @param extra - Metadata included in the entry; omitted or undefined adds no entry metadata.
   * @returns No value.
   */
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
 * @param handler - Synchronous custom sink; omitted or undefined uses the matching console method. Sink exceptions propagate; returned promises are not observed.
 * @returns Logger methods
 *
 * @example Console sink
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxLogger } from '@apvee/spfx-react-toolkit';
 *
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
 * @example Custom telemetry sink
 * ```tsx
 * import * as React from 'react';
 * import { LogEntry, useSPFxLogger } from '@apvee/spfx-react-toolkit';
 * function SaveButton({ publish }: { publish: (entry: LogEntry) => void }) {
 *   const logger = useSPFxLogger(publish);
 *   return <button onClick={() => logger.info('Save selected', { action: 'save' })}>Save</button>;
 * }
 * ```
 * @throws When called outside an SPFx provider.
 *
 * @useWhen application events need structured host, user and correlation fields.
 *
 * @scenario Write native log levels and inspect recent calls
 * ```tsx
 * // Requires the matching SPFx provider with native PageContext. Commands use the default console sink with synthetic messages; no host metadata is copied into UI history. A supplied region renderer owns collapsed diagnostics. Reset clears local UI only. No log is emitted on mount.
 * import * as React from 'react';
 * import { Button, makeStyles, tokens } from '@fluentui/react-components';
 * import { useSPFxContext, useSPFxLogger, useSx, padding } from '@apvee/spfx-react-toolkit';
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
 * const useLogStyles = makeStyles({
 *   message: { marginBottom: tokens.spacingVerticalXS, borderRadius: tokens.borderRadiusMedium, fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
 *   info: { backgroundColor: tokens.colorNeutralBackground2 },
 *   warning: { backgroundColor: tokens.colorStatusWarningBackground1 },
 *   error: { backgroundColor: tokens.colorStatusDangerBackground1 },
 * });
 * type Call = { level: 'info' | 'warning' | 'error'; message: string };
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const { spfxContext } = useSPFxContext();
 *   const logger = useSPFxLogger();
 *   const sx = useSx();
 *   const styles = useLogStyles();
 *   const [calls, setCalls] = React.useState<Call[]>([]);
 *   const [failure, setFailure] = React.useState<string>();
 *   const lifetime = React.useMemo(() => ({ active: true }), [resetRevision, spfxContext, logger]);
 *   React.useLayoutEffect(() => {
 *     setCalls([]);
 *     setFailure(undefined);
 *     return () => { lifetime.active = false; };
 *   }, [lifetime]);
 *   const write = (level: Call['level']): void => {
 *     if (!lifetime.active) return;
 *     const message = `SPFx toolkit demo ${level}`;
 *     try {
 *       if (level === 'warning') logger.warn(message);
 *       else logger[level](message);
 *       setFailure(undefined);
 *       setCalls(previous => [...previous.slice(-4), { level, message }]);
 *     } catch (error) {
 *       setFailure(error instanceof Error ? error.message : String(error));
 *     }
 *   };
 *   const latest = calls[calls.length - 1];
 *   return render({
 *     commands: <><Button onClick={() => write('info')}>Info</Button><Button onClick={() => write('warning')}>Warning</Button><Button onClick={() => write('error')}>Error</Button></>,
 *     status: failure ? <p role="alert">{failure}</p> : undefined,
 *     children: latest ? <div data-runtime-result className={sx(styles.message, styles[latest.level], padding.small)}><strong>{latest.level.toUpperCase()}</strong>: {latest.message}</div> : <p>Run a command to log a message.</p>,
 *     diagnostics: calls.length ? <ol aria-label="Recent log calls">{calls.map((call, index) => <li key={index}>{call.level.toUpperCase()}: {call.message}</li>)}</ol> : undefined,
 *   });
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Info, Warning and Error call the native console logger with safe synthetic demo messages; the current result replaces its predecessor and collapsed diagnostics retain the last five UI calls. Reset clears only this UI history, not console output.' };
 * ```
 */
export function useSPFxLogger(
  handler?: /**
    * Receives each structured log entry synchronously.
    * @param entry - Message with host/user/correlation metadata.
    * @returns No value; asynchronous completion is not observed.
    */ (entry: LogEntry) => void
): SPFxLoggerInfo {
  const { id: instanceId, kind } = useSPFxInstanceInfo();
  const { displayName, loginName } = useSPFxUserInfo();
  const { siteUrl, webUrl } = useSPFxSiteInfo();
  const { correlationId } = useSPFxCorrelationInfo();
  const { spfxContext } = useSPFxContext();
  
  const emit = useCallback((level: LogLevel, message: string, extra?: Record<string, unknown>): void => {
    // Extract webPartTag only if in WebPart context
    let webPartTag: string | undefined;
    if (kind === 'WebPart') {
      const wpContext = spfxContext as WebPartContext;
      webPartTag = wpContext.webPartTag;
    }
    
    const entry: LogEntry = {
      level,
      message,
      ts: new Date().toISOString(),
      instanceId,
      host: kind,
      user: displayName + ' (' + loginName + ')',
      siteUrl,
      webUrl,
      correlationId,
      webPartTag,
      extra,
    };
    
    if (handler) {
      handler(entry);
    } else {
      // Default: log to console
      const levelUpper = level.toUpperCase();
      // For WebPart contexts, display webPartTag instead of instanceId for better readability
      const displayId = kind === 'WebPart' && webPartTag ? webPartTag : instanceId;
      const line = '[' + levelUpper + '] ' + entry.ts + ' ' + kind + '/' + displayId + ' – ' + message;
      
      // Use appropriate console method
      const consoleFn = level === 'debug' ? console.log : console[level];
      consoleFn(line, extra ?? {});
    }
  }, [instanceId, kind, displayName, loginName, siteUrl, webUrl, correlationId, spfxContext, handler]);
  
  return useMemo(() => ({
    debug: (m: string, e?: Record<string, unknown>): void => emit('debug', m, e),
    info: (m: string, e?: Record<string, unknown>): void => emit('info', m, e),
    warn: (m: string, e?: Record<string, unknown>): void => emit('warn', m, e),
    error: (m: string, e?: Record<string, unknown>): void => emit('error', m, e),
  }), [emit]);
}
