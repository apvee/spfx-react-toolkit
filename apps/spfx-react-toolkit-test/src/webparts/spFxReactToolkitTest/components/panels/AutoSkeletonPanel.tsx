import * as React from 'react';
import { FluentProvider, webDarkTheme, webLightTheme, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { AutoSkeleton, AutoSkeletonProps, useSx, useStableCallback, flex, gap, padding, width } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

const useStyles = makeStyles({
  preview: { border: `2px solid ${tokens.colorNeutralStroke1}`, padding: '12px', boxSizing: 'border-box', maxWidth: '100%' },
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
    <label>Persistent note <input data-auto-skeleton-demo="focus-target" ref={inputRef} className={styles.visibleOverride}
      value={note} onChange={event => setNote(event.target.value)} /></label>
    <div data-skeleton-ignore=""><span>Ignored subtree: this text never creates an item.</span></div>
    <div className="auto-skeleton-demo-ignore"><span>Selector-controlled subtree.</span></div>
  </div>;
};

const AutoSkeletonExamples: React.FC = () => {
  const sx = useSx();
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
  return <DemoCard title="AutoSkeleton" iconName="Processing" description="Skeleton items follow mounted HTML geometry; loading keeps local child state.">
    <div className={sx(flex.column, gap.medium, padding.medium)}>
      <label><input data-auto-skeleton-demo="loading" type="checkbox" checked={loading} onChange={event => setLoading(event.target.checked)} /> Loading</label>
      <label>Animation <select data-auto-skeleton-demo="animation" value={animation} onChange={event => {
        const selected = animations.find(value => value === event.target.value); if (selected) setAnimation(selected);
      }}>{animations.map(value => <option key={value}>{value}</option>)}</select></label>
      <label>Minimum size <select data-auto-skeleton-demo="minSize" value={minSize} onChange={event => setMinSize(Number(event.target.value))}>
        {[8, 16, 32].map(value => <option key={value} value={value}>{value}px</option>)}
      </select></label>
      <label><input data-auto-skeleton-demo="ignoreSelectors" type="checkbox" checked={ignore} onChange={event => setIgnore(event.target.checked)} /> Ignore selector-controlled subtree</label>
      <label><input data-auto-skeleton-demo="resize" type="checkbox" checked={narrow} onChange={event => setNarrow(event.target.checked)} /> Narrow preview, including while loading</label>
      <label><input data-auto-skeleton-demo="empty" type="checkbox" checked={empty} onChange={event => setEmpty(event.target.checked)} /> Empty content</label>
      <button data-auto-skeleton-demo="focus-load" type="button" disabled={loading || empty} onClick={() => {
        inputRef.current?.focus(); setLoading(true);
      }}>Focus note, then load</button>
      <AutoSkeleton data-auto-skeleton-demo="preview" ref={rootRef} loading={loading} animation={animation} minSize={minSize}
        ignoreSelectors={ignore ? ignoredSelectors : noSelectors} className={mergeClasses(styles.preview, sx(width.px(narrow ? 220 : 360)))}
        onClick={() => setClicks(value => value + 1)}>
        {empty ? undefined : <MountedContent inputRef={inputRef} onMount={onMount} />}
      </AutoSkeleton>
      <button data-auto-skeleton-demo="inspect-ref" type="button" onClick={() => {
        const root = rootRef.current;
        setObserved(root ? `${root.tagName.toLowerCase()}, ${root.querySelectorAll('[data-auto-skeleton-shape]').length} items, busy=${root.getAttribute('aria-busy')}` : 'unmounted');
      }}>Inspect root ref</button>
      <p data-auto-skeleton-demo="observations">Root: {observed}. Content mounts: {mounts}. Preview clicks: {clicks}.</p>
      <p>Type a note, toggle loading, and return to loaded content: the value and mount count stay unchanged. Empty content has no inferred placeholder. Native inert support is required to hide content safely.</p>
    </div>
  </DemoCard>;
};

const AutoSkeletonPanel: React.FC = () => {
  const [dark, setDark] = React.useState(false);
  const [rtl, setRtl] = React.useState(false);
  return <>
    <label><input data-auto-skeleton-demo="theme" type="checkbox" checked={dark} onChange={event => setDark(event.target.checked)} /> Dark theme</label>{' '}
    <label><input data-auto-skeleton-demo="direction" type="checkbox" checked={rtl} onChange={event => setRtl(event.target.checked)} /> Right-to-left</label>
    <FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir={rtl ? 'rtl' : 'ltr'}><AutoSkeletonExamples /></FluentProvider>
  </>;
};
export default AutoSkeletonPanel;
