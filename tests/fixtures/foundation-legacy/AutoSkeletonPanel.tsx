/** Regression snapshot captured before single-subject consolidation; not product UI. */
import * as React from 'react';
import { FluentProvider, makeStyles, mergeClasses, tokens, Checkbox, Select, Button, Field, Input } from '@fluentui/react-components';
import { AutoSkeleton, AutoSkeletonProps, useSx, useStableCallback, flex, gap, width, background, minWidth } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';

const useStyles = makeStyles({
  preview: { boxSizing: 'border-box', maxWidth: '100%' },
  rectangle: { width: '160px', height: '24px', borderRadius: '4px', backgroundColor: tokens.colorNeutralBackground3 },
  circle: { width: '40px', height: '40px', borderRadius: '50%', backgroundColor: tokens.colorBrandBackground },
  visibleOverride: { visibility: 'visible' },
});
const ignoredSelectors: readonly string[] = ['.auto-skeleton-demo-ignore'];
const noSelectors: readonly string[] = [];
const animations: readonly NonNullable<AutoSkeletonProps['animation']>[] = ['wave', 'pulse'];

const MountedContent: React.FC<{ inputRef: React.RefObject<HTMLInputElement>; onMount: () => void }> = ({ inputRef, onMount }) => {
  const sx = useSx();
  const styles = useStyles();
  const [note, setNote] = React.useState('This value survives loading');
  React.useEffect(() => { onMount(); }, [onMount]);
  return <div data-skeleton-container="" className={sx(flex.column, gap.small)}>
    <div data-auto-skeleton-demo="circle" data-skeleton-shape="circle" className={styles.circle} />
    <div data-auto-skeleton-demo="rectangle" data-skeleton-shape="rectangle" className={styles.rectangle} />
    <p data-auto-skeleton-demo="mixed-text">A bounded profile with <strong>mixed inline text</strong> wraps into real measured lines as the preview width changes. Loading preserves its layout.</p>
    <Field label="Persistent note"><Input input={{ as: 'input', ref: inputRef, ...{ 'data-auto-skeleton-demo': 'focus-target' } }}  className={styles.visibleOverride}
      value={note} onChange={(_event, data) => setNote(data.value)} /></Field>
    <div data-skeleton-ignore=""><span>Ignored subtree: this text never creates an item.</span></div>
    <div className="auto-skeleton-demo-ignore"><span>Selector-controlled subtree.</span></div>
  </div>;
};

const AutoSkeletonExamples: React.FC<{ directionControl: React.ReactNode }> = ({ directionControl }) => {
  const sx = useSx();
  const demoStyles = useDemoStyles();
  const styles = useStyles();
  const [loading, setLoading] = React.useState(false);
  const [empty, setEmpty] = React.useState(false);
  const [narrow, setNarrow] = React.useState(false);
  const [ignore, setIgnore] = React.useState(true);
  const [animation, setAnimation] = React.useState<NonNullable<AutoSkeletonProps['animation']>>('wave');
  const [minSize, setMinSize] = React.useState(8);
  const [mounts, setMounts] = React.useState(0);
  const [clicks, setClicks] = React.useState(0);
  const [observed, setObserved] = React.useState('not inspected');
  const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const inputRef = React.useMemo(() => React.createRef<HTMLInputElement>(), []);
  const onMount = useStableCallback(() => setMounts(value => value + 1));
  return <DemoCard title="AutoSkeleton" iconName="Processing" description="Skeleton items follow mounted HTML geometry; loading keeps local child state."
    controls={<div className={demoStyles.controls}>
      {directionControl}
      <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'loading' } }} checked={loading} onChange={(_event, data) => setLoading(data.checked === true)} label="Loading" />
      <Field label="Animation"><Select select={{ as: 'select', ...{ 'data-auto-skeleton-demo': 'animation' } }} value={animation} onChange={(_event, data) => {
        const selected = animations.find(value => value === data.value); if (selected) setAnimation(selected);
      }}>{animations.map(value => <option key={value}>{value}</option>)}</Select></Field>
      <Field label="Minimum size"><Select select={{ as: 'select', ...{ 'data-auto-skeleton-demo': 'minSize' } }} value={minSize} onChange={(_event, data) => setMinSize(Number(data.value))}>
        {[8, 16, 32].map(value => <option key={value} value={value}>{value}px</option>)}
      </Select></Field>
      <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'ignoreSelectors' } }} checked={ignore} onChange={(_event, data) => setIgnore(data.checked === true)} label="Ignore selector-controlled subtree" />
      <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'resize' } }} checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} label="Narrow preview, including while loading" />
      <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'empty' } }} checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} label="Empty content" />

    </div>}
    commands={<>
      <Button data-auto-skeleton-demo="focus-load" type="button" disabled={loading || empty} onClick={() => {
        inputRef.current?.focus(); setLoading(true);
      }}>Focus note, then load</Button>
    </>}
    status={<p aria-live="polite">{loading ? 'Loading' : 'Loaded'} · Content mounts: {mounts} · Preview clicks: {clicks}</p>}
    diagnostics={<>
      <Button data-auto-skeleton-demo="inspect-ref" type="button" onClick={() => {
        const root = rootRef.current;
        setObserved(root ? `${root.tagName.toLowerCase()}, ${root.querySelectorAll('[data-auto-skeleton-shape]').length} items, busy=${root.getAttribute('aria-busy')}` : 'unmounted');
      }}>Inspect root ref</Button>
      <p data-auto-skeleton-demo="observations">Root: {observed}. Content mounts: {mounts}. Preview clicks: {clicks}.</p>
      <p>Type a note, toggle loading, and return to loaded content: the value and mount count stay unchanged. Empty content has no inferred placeholder. Native inert support is required to hide content safely.</p>

      </>}>
      <AutoSkeleton data-auto-skeleton-demo="preview" ref={rootRef} loading={loading} animation={animation} minSize={minSize}
        ignoreSelectors={ignore ? ignoredSelectors : noSelectors} className={mergeClasses(styles.preview, sx(width.px(narrow ? 220 : 360)))}
        onClick={() => setClicks(value => value + 1)}>
        {empty ? undefined : <MountedContent inputRef={inputRef} onMount={onMount} />}
      </AutoSkeleton>
  </DemoCard>;
};

const AutoSkeletonPanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  return <>

    <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><AutoSkeletonExamples directionControl={<Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'direction' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="Override host direction to RTL" />} /></FluentProvider>
  </>;
};
export default AutoSkeletonPanel;
