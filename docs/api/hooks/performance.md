# Performance & Diagnostics Hooks

> Hooks for performance monitoring, logging, and diagnostic information

## Overview

These hooks provide performance measurement, logging, correlation tracking, and tenant configuration access.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxPerformance`](#usespfxperformance) | `SPFxPerformanceResult` | Performance timing utilities |
| [`useSPFxLogger`](#usespfxlogger) | `SPFxLoggerResult` | Structured logging |
| [`useSPFxCorrelationInfo`](#usespfxcorrelationinfo) | `SPFxCorrelationInfo` | Request correlation IDs |
| [`useSPFxTenantProperty`](#usespfxtenantproperty) | `SPFxTenantPropertyResult` | Tenant-wide properties |

---

## useSPFxPerformance

Performance measurement and timing utilities.

### Signature

```typescript
function useSPFxPerformance(): SPFxPerformanceResult
```

### Returns

```typescript
interface PerformanceMark {
  /** Mark name */
  readonly name: string;
  
  /** Timestamp in milliseconds */
  readonly timestamp: number;
}

interface PerformanceMeasure {
  /** Measure name */
  readonly name: string;
  
  /** Duration in milliseconds */
  readonly duration: number;
  
  /** Start mark name */
  readonly startMark: string;
  
  /** End mark name */
  readonly endMark: string;
}

interface SPFxPerformanceResult {
  /**
   * Create a performance mark.
   * @param name - Unique mark name
   */
  readonly mark: (name: string) => void;
  
  /**
   * Measure time between two marks.
   * @param name - Measure name
   * @param startMark - Start mark name
   * @param endMark - End mark name (defaults to now)
   */
  readonly measure: (name: string, startMark: string, endMark?: string) => PerformanceMeasure;
  
  /**
   * Get all performance marks.
   */
  readonly getMarks: () => PerformanceMark[];
  
  /**
   * Get all performance measures.
   */
  readonly getMeasures: () => PerformanceMeasure[];
  
  /**
   * Clear all marks and measures.
   */
  readonly clear: () => void;
  
  /**
   * Time a function execution.
   * @param name - Measure name
   * @param fn - Function to time
   */
  readonly time: <T>(name: string, fn: () => T) => T;
  
  /**
   * Time an async function execution.
   * @param name - Measure name
   * @param fn - Async function to time
   */
  readonly timeAsync: <T>(name: string, fn: () => Promise<T>) => Promise<T>;
}
```

### Example: Measure Data Loading

```tsx
import { useSPFxPerformance } from '@apvee/spfx-react-toolkit';

function DataList() {
  const { mark, measure, timeAsync } = useSPFxPerformance();
  const [items, setItems] = React.useState<IItem[]>([]);
  
  React.useEffect(() => {
    const loadData = async () => {
      mark('data-load-start');
      
      const data = await timeAsync('fetch-items', () => 
        fetchItems()
      );
      
      mark('data-load-end');
      const loadTime = measure('total-load', 'data-load-start', 'data-load-end');
      
      console.log(`Data loaded in ${loadTime.duration}ms`);
      setItems(data);
    };
    
    loadData();
  }, []);
  
  return <ItemList items={items} />;
}
```

### Example: Performance Dashboard

```tsx
import { useSPFxPerformance } from '@apvee/spfx-react-toolkit';

function PerformanceMonitor() {
  const { getMeasures, clear } = useSPFxPerformance();
  const [measures, setMeasures] = React.useState<PerformanceMeasure[]>([]);
  
  React.useEffect(() => {
    const interval = setInterval(() => {
      setMeasures(getMeasures());
    }, 1000);
    
    return () => clearInterval(interval);
  }, []);
  
  return (
    <div className="perf-monitor">
      <h3>Performance Metrics</h3>
      <button onClick={clear}>Clear</button>
      <table>
        <thead>
          <tr>
            <th>Operation</th>
            <th>Duration</th>
          </tr>
        </thead>
        <tbody>
          {measures.map(m => (
            <tr key={m.name}>
              <td>{m.name}</td>
              <td>{m.duration.toFixed(2)}ms</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
```

### Example: Component Render Timing

```tsx
import { useSPFxPerformance } from '@apvee/spfx-react-toolkit';

function TrackedComponent() {
  const { mark, measure } = useSPFxPerformance();
  
  // Mark render start
  mark('render-start');
  
  // Effect runs after render
  React.useEffect(() => {
    mark('render-end');
    const renderTime = measure('component-render', 'render-start', 'render-end');
    
    if (renderTime.duration > 100) {
      console.warn(`Slow render: ${renderTime.duration}ms`);
    }
  });
  
  return <div>Tracked Content</div>;
}
```

### Source

[View source](../../src/hooks/useSPFxPerformance.ts)

---

## useSPFxLogger

Structured logging with SPFx Log service integration.

### Signature

```typescript
function useSPFxLogger(): SPFxLoggerResult
```

### Returns

```typescript
type LogLevel = 'verbose' | 'info' | 'warning' | 'error';

interface LogEntry {
  readonly timestamp: Date;
  readonly level: LogLevel;
  readonly source: string;
  readonly message: string;
  readonly data?: unknown;
}

interface SPFxLoggerResult {
  /**
   * Log verbose message (debug).
   */
  readonly verbose: (source: string, message: string, data?: unknown) => void;
  
  /**
   * Log info message.
   */
  readonly info: (source: string, message: string, data?: unknown) => void;
  
  /**
   * Log warning message.
   */
  readonly warn: (source: string, message: string, data?: unknown) => void;
  
  /**
   * Log error message.
   */
  readonly error: (source: string, message: string, error?: Error) => void;
  
  /**
   * Get recent log entries.
   */
  readonly getEntries: (count?: number) => LogEntry[];
  
  /**
   * Clear log entries.
   */
  readonly clear: () => void;
}
```

### Example: Service Logging

```tsx
import { useSPFxLogger } from '@apvee/spfx-react-toolkit';

function DataService() {
  const logger = useSPFxLogger();
  
  const fetchItems = async () => {
    const source = 'DataService.fetchItems';
    
    logger.info(source, 'Starting data fetch');
    
    try {
      const response = await fetch('/api/items');
      
      if (!response.ok) {
        logger.warn(source, 'Non-OK response', { status: response.status });
        throw new Error(`HTTP ${response.status}`);
      }
      
      const data = await response.json();
      logger.info(source, 'Data fetch complete', { count: data.length });
      
      return data;
    } catch (err) {
      logger.error(source, 'Data fetch failed', err as Error);
      throw err;
    }
  };
  
  return { fetchItems };
}
```

### Example: Debug Panel

```tsx
import { useSPFxLogger, useSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';

function DebugPanel() {
  const logger = useSPFxLogger();
  const { isLocal } = useSPFxEnvironmentInfo();
  const [entries, setEntries] = React.useState<LogEntry[]>([]);
  
  React.useEffect(() => {
    const interval = setInterval(() => {
      setEntries(logger.getEntries(50));
    }, 500);
    
    return () => clearInterval(interval);
  }, []);
  
  // Only show in local development
  if (!isLocal) return null;
  
  return (
    <div className="debug-panel">
      <h4>Debug Log</h4>
      <button onClick={logger.clear}>Clear</button>
      <div className="log-entries">
        {entries.map((entry, i) => (
          <div key={i} className={`log-entry log-${entry.level}`}>
            <span className="timestamp">
              {entry.timestamp.toLocaleTimeString()}
            </span>
            <span className="level">{entry.level.toUpperCase()}</span>
            <span className="source">{entry.source}</span>
            <span className="message">{entry.message}</span>
          </div>
        ))}
      </div>
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxLogger.ts)

---

## useSPFxCorrelationInfo

Access request correlation IDs for tracing.

### Signature

```typescript
function useSPFxCorrelationInfo(): SPFxCorrelationInfo
```

### Returns

```typescript
interface SPFxCorrelationInfo {
  /** Current correlation ID (GUID) */
  readonly correlationId: string;
  
  /** Session ID for the current page session */
  readonly sessionId: string;
  
  /** Generate a new correlation ID for a sub-request */
  readonly generateSubCorrelationId: () => string;
}
```

### Example: Request Tracing

```tsx
import { useSPFxCorrelationInfo, useSPFxLogger } from '@apvee/spfx-react-toolkit';

function TracedApiClient() {
  const { correlationId, generateSubCorrelationId } = useSPFxCorrelationInfo();
  const logger = useSPFxLogger();
  
  const fetchWithTracing = async (url: string) => {
    const subCorrelationId = generateSubCorrelationId();
    
    logger.info('API', `Request to ${url}`, { 
      correlationId, 
      subCorrelationId 
    });
    
    const response = await fetch(url, {
      headers: {
        'X-Correlation-Id': subCorrelationId,
        'X-Session-Id': correlationId
      }
    });
    
    logger.info('API', `Response from ${url}`, { 
      status: response.status,
      subCorrelationId 
    });
    
    return response;
  };
  
  return { fetchWithTracing };
}
```

### Example: Error Reporting

```tsx
import { useSPFxCorrelationInfo } from '@apvee/spfx-react-toolkit';

function ErrorBoundary({ children }: { children: React.ReactNode }) {
  const { correlationId, sessionId } = useSPFxCorrelationInfo();
  const [error, setError] = React.useState<Error | null>(null);
  
  const reportError = async (err: Error) => {
    await fetch('/api/errors', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: err.message,
        stack: err.stack,
        correlationId,
        sessionId,
        timestamp: new Date().toISOString()
      })
    });
  };
  
  React.useEffect(() => {
    if (error) {
      reportError(error);
    }
  }, [error]);
  
  if (error) {
    return (
      <div className="error-display">
        <h2>Something went wrong</h2>
        <p>Error ID: {correlationId}</p>
        <p>Please reference this ID when contacting support.</p>
      </div>
    );
  }
  
  return <>{children}</>;
}
```

### Source

[View source](../../src/hooks/useSPFxCorrelationInfo.ts)

---

## useSPFxTenantProperty

Access tenant-wide custom properties.

### Signature

```typescript
function useSPFxTenantProperty(key: string): SPFxTenantPropertyResult
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `key` | `string` | Yes | Tenant property key |

### Returns

```typescript
interface SPFxTenantPropertyResult {
  /** Property value (undefined while loading or if not found) */
  readonly value: string | undefined;
  
  /** Loading state */
  readonly isLoading: boolean;
  
  /** Error if fetch failed */
  readonly error: Error | undefined;
}
```

### Example: Feature Flags

```tsx
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';

function FeatureGatedComponent() {
  const { value: featureFlags, isLoading } = useSPFxTenantProperty('FeatureFlags');
  
  if (isLoading) return <Spinner />;
  
  const flags = featureFlags ? JSON.parse(featureFlags) : {};
  
  return (
    <div>
      {flags.enableNewUI && <NewUIComponent />}
      {flags.enableBetaFeatures && <BetaFeatures />}
    </div>
  );
}
```

### Example: Tenant Configuration

```tsx
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';

function TenantConfiguredComponent() {
  const { value: apiEndpoint, isLoading, error } = useSPFxTenantProperty('CustomApiEndpoint');
  const { value: apiKey } = useSPFxTenantProperty('CustomApiKey');
  
  if (isLoading) return <Spinner label="Loading configuration..." />;
  if (error) return <MessageBar messageBarType={MessageBarType.error}>{error.message}</MessageBar>;
  
  if (!apiEndpoint || !apiKey) {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        Tenant properties not configured. Contact your administrator.
      </MessageBar>
    );
  }
  
  return <ApiClient endpoint={apiEndpoint} apiKey={apiKey} />;
}
```

### Example: Multi-Property Configuration

```tsx
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';

function ConfiguredWidget() {
  const logo = useSPFxTenantProperty('CompanyLogo');
  const theme = useSPFxTenantProperty('CompanyTheme');
  const helpUrl = useSPFxTenantProperty('HelpDeskUrl');
  
  const isLoading = logo.isLoading || theme.isLoading || helpUrl.isLoading;
  
  if (isLoading) return <Spinner />;
  
  const themeColors = theme.value ? JSON.parse(theme.value) : { primary: '#0078d4' };
  
  return (
    <div style={{ '--primary-color': themeColors.primary } as React.CSSProperties}>
      {logo.value && <img src={logo.value} alt="Company Logo" />}
      {helpUrl.value && <a href={helpUrl.value}>Help</a>}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxTenantProperty.ts)

---

## See Also

- [Context Hooks](./context.md) - SPFx context access
- [Environment Hooks](./environment.md) - Environment detection
- [Storage Hooks](./storage.md) - Data persistence

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
