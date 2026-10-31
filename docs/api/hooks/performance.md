# Performance & Diagnostics Hooks

[React and SPFx hooks](../INDEX.md#react-and-spfx-hooks)

## useSPFxPerformance

```typescript
function useSPFxPerformance(): SPFxPerformanceInfo;
interface SPFxPerfResult<T = unknown> {
  readonly name: string;
  readonly durationMs: number;
  readonly result?: T;
  readonly instanceId: string;
  readonly host: string;
  readonly correlationId: string | undefined;
}
interface SPFxPerformanceInfo {
  readonly mark: (name: string) => void;
  readonly measure: (name: string, startMark: string, endMark?: string) => SPFxPerfResult;
  readonly time: <T>(name: string, fn: () => Promise<T> | T) => Promise<SPFxPerfResult<T>>;
}
```

`time()` awaits sync or async work and returns its result with measured duration and SPFx metadata. Concurrent calls with the same measurement name have distinct internal start marks, removed in `finally`; callback failures reject the returned promise. Browser measurement failures fall back to a zero duration. Public `mark()` names and measurement entries remain part of the browser Performance API; callers choosing explicit mark names manage collisions and cleanup themselves.

```tsx
import * as React from 'react';
import { useSPFxPerformance } from '@apvee/spfx-react-toolkit';

function TimedButton() {
  const { time } = useSPFxPerformance();
  const run = async () => {
    const measured = await time('calculate', () => 2 + 2);
    console.log(measured.result, measured.durationMs);
  };
  return <button onClick={run}>Measure</button>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPerformance.ts)

## useSPFxLogger

```typescript
type LogLevel = 'debug' | 'info' | 'warn' | 'error';
interface LogEntry {
  readonly level: LogLevel;
  readonly message: string;
  readonly ts: string;
  readonly instanceId: string;
  readonly host: string;
  readonly user: string;
  readonly siteUrl: string | undefined;
  readonly webUrl: string | undefined;
  readonly correlationId: string | undefined;
  readonly webPartTag?: string;
  readonly extra?: Record<string, unknown>;
}
interface SPFxLoggerInfo {
  readonly debug: (message: string, extra?: Record<string, unknown>) => void;
  readonly info: (message: string, extra?: Record<string, unknown>) => void;
  readonly warn: (message: string, extra?: Record<string, unknown>) => void;
  readonly error: (message: string, extra?: Record<string, unknown>) => void;
}
function useSPFxLogger(handler?: (entry: LogEntry) => void): SPFxLoggerInfo;
```

Without a handler the hook logs to console; a custom handler receives the structured entry. Application code owns persistence and any external telemetry transport.

```tsx
import * as React from 'react';
import { useSPFxLogger } from '@apvee/spfx-react-toolkit';

function SaveButton() {
  const logger = useSPFxLogger(entry => console.log(entry));
  return <button onClick={() => logger.info('Save requested', { source: 'toolbar' })}>Save</button>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxLogger.ts)

## useSPFxCorrelationInfo

```typescript
function useSPFxCorrelationInfo(): SPFxCorrelationInfo;
interface SPFxCorrelationInfo {
  readonly correlationId: string | undefined;
  readonly tenantId: string | undefined;
}
```

Maps optional correlation and tenant IDs exposed by PageContext. A missing ID remains undefined; this hook does not create a distributed tracing session.

```tsx
import * as React from 'react';
import { useSPFxCorrelationInfo } from '@apvee/spfx-react-toolkit';

function CorrelationStatus() {
  const { correlationId, tenantId } = useSPFxCorrelationInfo();
  return <p>{tenantId ?? 'No tenant ID'} / {correlationId ?? 'No correlation ID'}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxCorrelationInfo.ts)

## Diagnostic member contracts

`SPFxPerformanceInfo.mark(name)` attempts a native browser mark and ignores unsupported/failed marks. `SPFxPerformanceInfo.measure(name, startMark, endMark?)` creates the supplied end mark at the time of the call, then reads the latest named measurement; omitted end mark measures to the current time. Unsupported/failed measurements return duration 0. `SPFxPerformanceInfo.time(name, fn)` calls sync/async work and resolves its measured result; failures reject with the original callback failure.

`SPFxPerfResult.name` is the caller's measurement name, `SPFxPerfResult.durationMs` is elapsed milliseconds, and `SPFxPerfResult.result` is populated only by successful `time` (possibly undefined if the callback returns it). `SPFxPerfResult.instanceId`, `SPFxPerfResult.host` and `SPFxPerfResult.correlationId` are the current provider metadata and optional host correlation ID. `host` is the SPFx component kind, not Teams/Office classification.

`SPFxLoggerInfo.debug`, `info`, `warn` and `error` accept a required message and optional `Record<string, unknown>` extra metadata, return no value, and call the supplied handler synchronously without swallowing its failures. Omitted handler uses console (debug uses console.log). `LogEntry.level` is `LogLevel`, `LogEntry.message` the input, `LogEntry.ts` an ISO timestamp, `LogEntry.instanceId`/`LogEntry.host` the provider identity/kind, and `LogEntry.user` formats display name plus login. `LogEntry.siteUrl`, `webUrl` and `correlationId` carry available host metadata; `LogEntry.webPartTag` is native WebPart-only metadata and `LogEntry.extra` is the optional original record. No buffering, redaction or persistence is added.

`SPFxCorrelationInfo.correlationId` and `SPFxCorrelationInfo.tenantId` map available native IDs to strings or undefined. Both performance/logger hooks require a matching provider; they do not initialize telemetry sessions or publish data externally.

## See Also

- [Context Hooks](context.md)
- [Page Context Helpers](../helpers/page-context.md#page-context-helpers)
