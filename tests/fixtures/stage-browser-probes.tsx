import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';
import { Button, FluentProvider, makeStyles, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { Stage } from '@apvee/spfx-react-toolkit';

interface ProbeState {
  dark: boolean;
  trigger: boolean;
  longActions: boolean;
  preventEscape: boolean;
  preventTrigger: boolean;
  visible: boolean;
  retain: boolean;
}
interface ProbeObservations { titleRef: boolean; descriptionRef: boolean; clicks: number; requests: number; mounts: number; }
declare global {
  interface Window {
    stageProbe: {
      configure: (patch: Partial<ProbeState>) => void;
      observe: () => ProbeObservations;
    };
  }
}
const useStyles = makeStyles({
  viewport: { position: 'fixed', top: 0, left: 0, width: '100vw', height: '100vh', visibility: 'hidden', pointerEvents: 'none' },
  region: { marginBlock: '16px' },
  frame: { width: '1000px', maxWidth: '100%', height: '800px' },
  title: { color: 'rgb(130, 40, 180)', fontSize: '18px' },
  description: { fontStyle: 'italic', letterSpacing: '1px' },
});
const ProbeCounter: React.FC<{ onMount: () => void }> = ({ onMount }) => {
  const [count, setCount] = React.useState(0);
  React.useEffect(onMount, [onMount]);
  return <><Button data-stage-probe="increment" onClick={() => setCount(value => value + 1)}>Probe increment</Button><output data-stage-probe="counter">{count}</output></>;
};

/** Supplementary native controls cover contracts that the unchanged demo cannot configure. */
const Probe: React.FC<{ targetDocument?: Document; iframe?: boolean }> = ({ targetDocument, iframe = false }) => {
  const styles = useStyles();
  const [state, setState] = React.useState<ProbeState>({ dark: false, trigger: true, longActions: false, preventEscape: false, preventTrigger: false, visible: true, retain: false });
  const [open, setOpen] = React.useState(false);
  const [clicks, setClicks] = React.useState(0);
  const [requests, setRequests] = React.useState(0);
  const [mounts, setMounts] = React.useState(0);
  const titleRef = React.useRef<HTMLElement>();
  const descriptionRef = React.useRef<HTMLSpanElement>();
  const onMount = React.useCallback(() => setMounts(value => value + 1), []);
  React.useEffect(() => {
    if (iframe) return;
    window.stageProbe = { configure: patch => setState(value => ({ ...value, ...patch })), observe: () => ({ titleRef: Boolean(titleRef.current?.isConnected), descriptionRef: Boolean(descriptionRef.current?.isConnected), clicks, requests, mounts }) };
  }, [iframe, clicks, requests, mounts]);
  return <FluentProvider theme={state.dark || iframe ? webDarkTheme : webLightTheme} dir={iframe ? 'rtl' : 'ltr'} targetDocument={targetDocument}>
    <section data-stage-probe={iframe ? 'iframe-region' : 'region'} className={styles.region}>
      <Button data-stage-probe="fallback">Probe fallback focus</Button>
      <label><input data-stage-probe="retention" type="checkbox" checked={state.retain}
        onChange={event => setState(value => ({ ...value, retain: event.target.checked }))} />Retain probe surface</label>
      {state.visible && <Stage presentation="dialog" size="small" open={open} unmountOnClose={!state.retain}
        trigger={state.trigger ? <Button data-stage-probe="trigger" onClick={event => { if (state.preventTrigger) event.preventDefault(); }}>Open probe</Button> : undefined}
        onOpenChange={(_, data) => { setRequests(value => value + 1); setOpen(data.open); }} data-stage-probe="surface">
        <Stage.Header title={<em data-stage-probe="jsx-title" className={styles.title}
          // React 17 callback refs receive null on detach; store undefined for absence.
          ref={element => { titleRef.current = element ?? undefined; }} onClick={() => setClicks(value => value + 1)}>Authored JSX title</em>}
          description={<span data-stage-probe="jsx-description" className={styles.description}
            ref={element => { descriptionRef.current = element ?? undefined; }}>Authored JSX description</span>}
          closeLabel={iframe ? 'Close iframe probe' : 'Close probe'}
          actions={state.longActions ? <><Button>First longer Header action text</Button><Button>Second longer Header action text</Button></> : <Button>Probe Header action</Button>} />
        <Stage.Body><input aria-label="Probe input" onKeyDown={event => { if (state.preventEscape && event.key === 'Escape') event.preventDefault(); }} />
          <ProbeCounter onMount={onMount} /></Stage.Body>
        <Stage.Footer><Button onClick={() => setOpen(false)}>Probe Footer close</Button></Stage.Footer>
      </Stage>}
      <output data-stage-probe="requests">{requests}</output>
      <output data-stage-probe="jsx-clicks">{clicks}</output>
      <output data-stage-probe="open">{String(open)}</output>
    </section>
  </FluentProvider>;
};

export const StageBrowserProbes: React.FC = () => {
  const styles = useStyles();
  const [frameDocument, setFrameDocument] = React.useState<Document>();
  const frameRenderer = React.useMemo(() => frameDocument ? createDOMRenderer(frameDocument) : undefined, [frameDocument]);
  return <>
    <div data-stage-probe="viewport" aria-hidden="true" className={styles.viewport} />
    <Probe />
    <iframe title="Independent Stage owner document" data-stage-probe="iframe" className={styles.frame}
      onLoad={event => setFrameDocument(event.currentTarget.contentDocument ?? undefined)} srcDoc="<!doctype html><html><head><title>Stage iframe</title></head><body></body></html>" />
    {frameDocument && frameRenderer && ReactDOM.createPortal(<RendererProvider renderer={frameRenderer} targetDocument={frameDocument}>
      <Probe targetDocument={frameDocument} iframe />
    </RendererProvider>, frameDocument.body)}
  </>;
};
