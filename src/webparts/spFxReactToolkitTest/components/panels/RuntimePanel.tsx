import * as React from 'react';
import {
  DefaultButton,
  Label,
  PrimaryButton,
  Stack,
  TextField,
} from '@fluentui/react';
import {
  useSPFxLocalStorage,
  useSPFxLogger,
  useSPFxPerformance,
  useSPFxProperties,
  useSPFxSessionStorage,
} from '../../../../hooks';
import {
  ActionResult,
  DemoCard,
  InfoGrid,
} from '../shared';
import styles from '../SpFxReactToolkitTest.module.scss';

interface IWebPartProps {
  description?: string;
}

const RuntimePanel: React.FC = () => {
  const { properties, setProperties, updateProperties } = useSPFxProperties<IWebPartProps>();
  const sessionStorage = useSPFxSessionStorage('demo-session-key', '');
  const localStorage = useSPFxLocalStorage('demo-local-key', '');
  const logger = useSPFxLogger();
  const performance = useSPFxPerformance();

  const [description, setDescription] = React.useState(properties?.description ?? '');
  const [sessionInput, setSessionInput] = React.useState('');
  const [localInput, setLocalInput] = React.useState('');
  const [result, setResult] = React.useState<string>();
  const [perfResult, setPerfResult] = React.useState<string>();
  const [logs, setLogs] = React.useState<Array<{ level: string; message: string }>>([]);

  React.useEffect(() => {
    setDescription(properties?.description ?? '');
  }, [properties?.description]);

  const showResult = React.useCallback((message: string) => {
    setResult(message);
    window.setTimeout(() => setResult(undefined), 3000);
  }, []);

  const saveDescription = React.useCallback(() => {
    setProperties({ description });
    showResult('Properties updated with setProperties().');
  }, [description, setProperties, showResult]);

  const appendDescription = React.useCallback(() => {
    updateProperties(current => ({
      ...(current ?? {}),
      description: `${current?.description ?? ''} updated`.trim(),
    }));
    showResult('Properties updated with updateProperties().');
  }, [updateProperties, showResult]);

  const saveSession = React.useCallback(() => {
    sessionStorage.setValue(sessionInput);
    setSessionInput('');
    showResult('Session storage value saved.');
  }, [sessionInput, sessionStorage, showResult]);

  const saveLocal = React.useCallback(() => {
    localStorage.setValue(localInput);
    setLocalInput('');
    showResult('Local storage value saved.');
  }, [localInput, localStorage, showResult]);

  const runPerformance = React.useCallback(async () => {
    const timed = await performance.time('spfx-toolkit-demo', async () => {
      await new Promise(resolve => window.setTimeout(resolve, 250));
      return 'Timed operation completed';
    });

    setPerfResult(`${timed.result} in ${timed.durationMs.toFixed(2)}ms`);
  }, [performance]);

  const writeLog = React.useCallback((level: 'info' | 'warning' | 'error') => {
    const message = `SPFx toolkit demo ${level}`;

    if (level === 'info') {
      logger.info(message);
    } else if (level === 'warning') {
      logger.warn(message);
    } else {
      logger.error(message);
    }

    setLogs(previous => [...previous.slice(-4), { level, message }]);
  }, [logger]);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <ActionResult result={result} />

      <DemoCard title="Properties" iconName="Settings">
        <InfoGrid rows={[{ label: 'Current Description', value: properties?.description, icon: 'TextDocument' }]} />
        <TextField
          label="Description"
          value={description}
          onChange={(_, value) => setDescription(value ?? '')}
        />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="setProperties" iconProps={{ iconName: 'Save' }} onClick={saveDescription} />
          <DefaultButton text="updateProperties" iconProps={{ iconName: 'Edit' }} onClick={appendDescription} />
        </Stack>
      </DemoCard>

      <DemoCard title="Scoped Storage" iconName="Database">
        <InfoGrid
          rows={[
            { label: 'Session Value', value: String(sessionStorage.value || ''), icon: 'TemporaryUser' },
            { label: 'Local Value', value: String(localStorage.value || ''), icon: 'Save' },
          ]}
        />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <TextField
            value={sessionInput}
            onChange={(_, value) => setSessionInput(value ?? '')}
            placeholder="Session value"
            styles={{ root: { minWidth: 220 } }}
          />
          <PrimaryButton text="Save Session" onClick={saveSession} disabled={!sessionInput} />
          <DefaultButton text="Clear Session" onClick={sessionStorage.remove} />
        </Stack>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <TextField
            value={localInput}
            onChange={(_, value) => setLocalInput(value ?? '')}
            placeholder="Local value"
            styles={{ root: { minWidth: 220 } }}
          />
          <PrimaryButton text="Save Local" onClick={saveLocal} disabled={!localInput} />
          <DefaultButton text="Clear Local" onClick={localStorage.remove} />
        </Stack>
      </DemoCard>

      <DemoCard title="Logger" iconName="Code">
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Info" iconProps={{ iconName: 'Info' }} onClick={() => writeLog('info')} />
          <DefaultButton text="Warning" iconProps={{ iconName: 'Warning' }} onClick={() => writeLog('warning')} />
          <DefaultButton text="Error" iconProps={{ iconName: 'ErrorBadge' }} onClick={() => writeLog('error')} />
        </Stack>
        {logs.length > 0 && (
          <Stack tokens={{ childrenGap: 4 }}>
            <Label>Recent log calls</Label>
            {logs.map((log, index) => {
              const levelClass = log.level === 'info'
                ? styles.info
                : log.level === 'warning'
                  ? styles.warning
                  : styles.error;

              return (
                <div key={`${log.level}-${index}`} className={`${styles.logMessage} ${levelClass}`}>
                  <strong>{log.level.toUpperCase()}</strong>: {log.message}
                </div>
              );
            })}
          </Stack>
        )}
      </DemoCard>

      <DemoCard title="Performance" iconName="SpeedHigh">
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Run timed operation" iconProps={{ iconName: 'LightningBolt' }} onClick={runPerformance} />
        </Stack>
        <InfoGrid rows={[{ label: 'Last Result', value: perfResult, icon: 'Timer' }]} />
      </DemoCard>
    </Stack>
  );
};

export default RuntimePanel;
