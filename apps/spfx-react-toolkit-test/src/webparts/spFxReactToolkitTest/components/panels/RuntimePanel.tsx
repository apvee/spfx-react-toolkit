import { useDemoStyles } from '../shared/demoStyles';
import { Select, Button, Field, Input, Label } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxLocalStorage, useSPFxLogger, useSPFxPerformance, useSPFxProperties, useSPFxSessionStorage, } from '@apvee/spfx-react-toolkit';
import { ActionResult, DemoCard, InfoGrid, } from '../shared';
import styles from '../SpFxReactToolkitTest.module.scss';
interface IWebPartProps {
    description?: string;
}
const RuntimePanelPropertiesScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const { properties, setProperties, updateProperties } = useSPFxProperties<IWebPartProps>();
    const [description, setDescription] = React.useState(properties?.description ?? '');
    const [result, setResult] = React.useState<{
        owner: 'properties' | 'storage';
        message: string;
    }>();
    const resultTimer = React.useRef<number>();
    React.useEffect(() => () => { if (resultTimer.current !== undefined)
        window.clearTimeout(resultTimer.current); }, []);
    React.useEffect(() => {
        setDescription(properties?.description ?? '');
    }, [properties?.description]);
    const showResult = React.useCallback((owner: 'properties' | 'storage', message: string) => {
        if (resultTimer.current !== undefined)
            window.clearTimeout(resultTimer.current);
        setResult({ owner, message });
        resultTimer.current = window.setTimeout(() => {
            resultTimer.current = undefined;
            setResult(undefined);
        }, 3000);
    }, []);
    const saveDescription = React.useCallback(() => {
        setProperties({ description });
        showResult('properties', 'Properties updated with setProperties().');
    }, [description, setProperties, showResult]);
    const appendDescription = React.useCallback(() => {
        updateProperties(current => ({
            ...(current ?? {}),
            description: `${current?.description ?? ''} updated`.trim(),
        }));
        showResult('properties', 'Properties updated with updateProperties().');
    }, [updateProperties, showResult]);
    return <DemoCard title="Properties" status={result?.owner === 'properties' ? <ActionResult result={result.message}/> : undefined} description="Update the web part description and observe the current value." iconName="Settings" controls={<>{selector(false)}<>
      <Field label={"Description"}>
        <Input value={description} onChange={(_, data) => setDescription(data.value ?? '')}/>
      </Field>

    </></>} contentPlacement="configuration" commands={<>
      <Button appearance="primary" onClick={saveDescription}>{"setProperties"}</Button>
      <Button onClick={appendDescription}>{"updateProperties"}</Button>
    </>}>
      <InfoGrid rows={[{ label: 'Current Description', value: properties?.description, icon: 'TextDocument' }]}/>

    </DemoCard>;
};
const RuntimePanelStorageScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const sessionStorage = useSPFxSessionStorage('demo-session-key', '');
    const localStorage = useSPFxLocalStorage('demo-local-key', '');
    const [sessionInput, setSessionInput] = React.useState('');
    const [localInput, setLocalInput] = React.useState('');
    const [result, setResult] = React.useState<{
        owner: 'properties' | 'storage';
        message: string;
    }>();
    const resultTimer = React.useRef<number>();
    React.useEffect(() => () => { if (resultTimer.current !== undefined)
        window.clearTimeout(resultTimer.current); }, []);
    const showResult = React.useCallback((owner: 'properties' | 'storage', message: string) => {
        if (resultTimer.current !== undefined)
            window.clearTimeout(resultTimer.current);
        setResult({ owner, message });
        resultTimer.current = window.setTimeout(() => {
            resultTimer.current = undefined;
            setResult(undefined);
        }, 3000);
    }, []);
    const saveSession = React.useCallback(() => {
        sessionStorage.setValue(sessionInput);
        setSessionInput('');
        showResult('storage', 'Session storage value saved.');
    }, [sessionInput, sessionStorage, showResult]);
    const saveLocal = React.useCallback(() => {
        localStorage.setValue(localInput);
        setLocalInput('');
        showResult('storage', 'Local storage value saved.');
    }, [localInput, localStorage, showResult]);
    return <DemoCard title="Scoped Storage" status={result?.owner === 'storage' ? <ActionResult result={result.message}/> : undefined} description="Save and clear values in scoped session and local storage." iconName="Database" controls={<>{selector(false)}<>
      <div className={demoStyles.actions}>
        <Input aria-label={"Session value"} value={sessionInput} onChange={(_, data) => setSessionInput(data.value ?? '')} placeholder="Session value"/>

      </div>
      <div className={demoStyles.actions}>
        <Input aria-label={"Local value"} value={localInput} onChange={(_, data) => setLocalInput(data.value ?? '')} placeholder="Local value"/>

      </div>
    </></>} contentPlacement="configuration" commands={<>
      <Button appearance="primary" onClick={saveSession} disabled={!sessionInput}>{"Save Session"}</Button>
      <Button onClick={sessionStorage.remove}>{"Clear Session"}</Button>
      <Button appearance="primary" onClick={saveLocal} disabled={!localInput}>{"Save Local"}</Button>
      <Button onClick={localStorage.remove}>{"Clear Local"}</Button>
    </>}>
      <InfoGrid rows={[
            { label: 'Session Value', value: String(sessionStorage.value || ''), icon: 'TemporaryUser' },
            { label: 'Local Value', value: String(localStorage.value || ''), icon: 'Save' },
        ]}/>

    </DemoCard>;
};
const RuntimePanelLoggerScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const logger = useSPFxLogger();
    const [logs, setLogs] = React.useState<Array<{
        level: string;
        message: string;
    }>>([]);
    const writeLog = React.useCallback((level: 'info' | 'warning' | 'error') => {
        const message = `SPFx toolkit demo ${level}`;
        if (level === 'info') {
            logger.info(message);
        }
        else if (level === 'warning') {
            logger.warn(message);
        }
        else {
            logger.error(message);
        }
        setLogs(previous => [...previous.slice(-4), { level, message }]);
    }, [logger]);
    const latest = logs[logs.length - 1];
    return <DemoCard controls={selector(false)} title="Logger" description="Write a log message and inspect the current call." iconName="Code" contentPlacement="configuration" commands={<>
      <Button onClick={() => writeLog('info')}>Info</Button>
      <Button onClick={() => writeLog('warning')}>Warning</Button>
      <Button onClick={() => writeLog('error')}>Error</Button>
    </>} diagnostics={logs.length ? <div className={demoStyles.panel}><Label>Recent log calls</Label>{logs.map((log, index) => <div key={`${log.level}-${index}`}>{log.level.toUpperCase()}: {log.message}</div>)}</div> : undefined}>
      {latest ? <div className={`${styles.logMessage} ${latest.level === 'info' ? styles.info : latest.level === 'warning' ? styles.warning : styles.error}`}><strong>{latest.level.toUpperCase()}</strong>: {latest.message}</div> : <Label>Run a command to log a message.</Label>}
    </DemoCard>;
};
const RuntimePanelPerformanceScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const performance = useSPFxPerformance();
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const [perfResult, setPerfResult] = React.useState<string>();
    const [duration, setDuration] = React.useState<string>();
    const runPerformance = React.useCallback(async () => {
        const timed = await performance.time('spfx-toolkit-demo', async () => {
            await new Promise(resolve => window.setTimeout(resolve, 250));
            return 'Timed operation completed';
        });
        if (mounted.current) {
            setPerfResult(timed.result);
            setDuration(`${timed.durationMs.toFixed(2)}ms`);
        }
    }, [performance]);
    return <DemoCard controls={selector(false)} title="Performance" diagnostics={<InfoGrid rows={[{ label: 'Duration', value: duration }]} />} description="Measure a timed asynchronous operation." iconName="SpeedHigh" contentPlacement="configuration" commands={<>
      <Button appearance="primary" onClick={runPerformance}>{"Run timed operation"}</Button>
    </>}>
      <InfoGrid rows={[{ label: 'Last Result', value: perfResult, icon: 'Timer' }]}/>
    </DemoCard>;
};
const RuntimePanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('properties');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'RuntimePanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}><option value="properties">Properties</option><option value="storage">Local storage</option><option value="logger">Logger</option><option value="performance">Performance</option></Select></Field>;
    return <>{scenario === 'properties' && <RuntimePanelPropertiesScenario selector={selector}/>}{scenario === 'storage' && <RuntimePanelStorageScenario selector={selector}/>}{scenario === 'logger' && <RuntimePanelLoggerScenario selector={selector}/>}{scenario === 'performance' && <RuntimePanelPerformanceScenario selector={selector}/>}</>;
};
export default RuntimePanel;
