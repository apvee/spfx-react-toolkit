import * as React from 'react';
import { Avatar, Card, Text, FluentProvider, makeStyles, mergeClasses, tokens, Checkbox, Select, Button, Field, Input } from '@fluentui/react-components';
import { Calendar16Regular, Person16Regular } from '@fluentui/react-icons';
import { AutoSkeleton, AutoSkeletonProps, useSx, useStableCallback, alignItems, flex, gap, width, background, minWidth } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';

const useStyles = makeStyles({
  preview: { boxSizing: 'border-box', maxWidth: '100%' },
  rectangle: { width: '160px', height: '24px', borderRadius: '4px', maxWidth: '100%' },
  circle: { width: '40px', height: '40px', flexShrink: 0 },
  secondary: { color: tokens.colorNeutralForeground2 },
  paragraph: { margin: 0 },
  visibleOverride: { visibility: 'visible' },
});
const ignoredSelectors: readonly string[] = ['.auto-skeleton-demo-ignore'];
const noSelectors: readonly string[] = [];
const animations: readonly NonNullable<AutoSkeletonProps['animation']>[] = ['wave', 'pulse'];

const MountedContent: React.FC<{ inputRef: React.RefObject<HTMLInputElement>; onMount: () => void }> = ({ inputRef, onMount }) => {
  const sx = useSx();
  const styles = useStyles();
  const [note, setNote] = React.useState('Documentation update');
  React.useEffect(() => { onMount(); }, [onMount]);
  return <div data-skeleton-container="" className={sx(flex.column, gap.medium)}>
    <div data-skeleton-container="" className={sx(flex.row, alignItems.center, gap.medium, minWidth.zero)}>
      <Avatar name="Giulia Rossi" size={40} data-auto-skeleton-demo="circle" data-skeleton-shape="circle" className={styles.circle} />
      <div data-skeleton-container="" className={sx(flex.column, minWidth.zero)}>
        <Text block weight="semibold" size={400} data-auto-skeleton-demo="rectangle" data-skeleton-shape="rectangle" className={styles.rectangle}>Giulia Rossi</Text>
        <Text size={200} className={styles.secondary}>Project coordinator</Text>
      </div>
    </div>
    <Text as="p" data-auto-skeleton-demo="mixed-text" className={styles.paragraph}>Giulia Rossi coordinates the <strong>intranet project</strong>. The team is completing the documentation for the next review.</Text>
    <Field label="Project note"><Input input={{ as: 'input', ref: inputRef, ...{ 'data-auto-skeleton-demo': 'focus-target' } }}  className={styles.visibleOverride}
      value={note} onChange={(_event, data) => setNote(data.value)} /></Field>
    <div data-skeleton-ignore="" className={sx(flex.row, alignItems.center, gap.small)}><Person16Regular aria-hidden="true" /><Text size={200}>Team lead: Giulia Rossi.</Text></div>
    <div className={mergeClasses('auto-skeleton-demo-ignore', sx(flex.row, alignItems.center, gap.small))}><Calendar16Regular aria-hidden="true" /><Text size={200}>Next review: Friday.</Text></div>
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
  return <DemoCard title="AutoSkeleton" iconName="Processing" description="Show the loading state while preserving profile content, note and layout."
    controls={<><Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'loading' } }} checked={loading} onChange={(_event, data) => setLoading(data.checked === true)} label="Loading" />
<Field label="Animation"><Select select={{ as: 'select', ...{ 'data-auto-skeleton-demo': 'animation', 'data-demo-variant': 'auto-skeleton' } }} value={animation} onChange={(_event, data) => {
        const selected = animations.find(value => value === data.value); if (selected) setAnimation(selected);
      }}>{animations.map(value => <option key={value}>{value}</option>)}</Select></Field>
<Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'resize' } }} checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} label="Narrow width" />
<Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'empty' } }} checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} label="Empty content" /></>}
    commands={<>
      <Button data-auto-skeleton-demo="focus-load" type="button" disabled={loading || empty} onClick={() => {
        inputRef.current?.focus(); setLoading(true);
      }}>Focus the note and load</Button>
    </>}
    status={<p aria-live="polite">{loading ? 'Loading' : 'Loaded'} · Content mounts: {mounts} · Preview clicks: {clicks}</p>}
    diagnostics={<><div className={demoStyles.controls}>{directionControl}
<Field label="Minimum size"><Select select={{ as: 'select', ...{ 'data-auto-skeleton-demo': 'minSize' } }} value={minSize} onChange={(_event, data) => setMinSize(Number(data.value))}>
        {[8, 16, 32].map(value => <option key={value} value={value}>{value}px</option>)}
      </Select></Field>
<Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'ignoreSelectors' } }} checked={ignore} onChange={(_event, data) => setIgnore(data.checked === true)} label="Ignore selected content" /></div>
<Button data-auto-skeleton-demo="inspect-ref" type="button" onClick={() => {
        const root = rootRef.current;
        setObserved(root ? `${root.tagName.toLowerCase()}, ${root.querySelectorAll('[data-auto-skeleton-shape]').length} items, busy=${root.getAttribute('aria-busy')}` : 'unmounted');
      }}>Inspect root ref</Button>

<p data-auto-skeleton-demo="observations">Root: {observed}. Content mounts: {mounts}. Preview clicks: {clicks}.</p>

<p>Type a note, toggle loading, and return to loaded content: the value and mount count stay unchanged. Empty content has no inferred placeholder. Native inert support is required to hide content safely.</p>
</>}>
      <Card data-auto-skeleton-demo="card" appearance="filled" size="medium" focusMode="off" className={mergeClasses(styles.preview, sx(width.px(narrow ? 220 : 360)))}>
        <AutoSkeleton data-auto-skeleton-demo="preview" ref={rootRef} loading={loading} animation={animation} minSize={minSize}
          ignoreSelectors={ignore ? ignoredSelectors : noSelectors} className={sx(width.full, minWidth.zero)}
          onClick={() => setClicks(value => value + 1)}>
          {empty ? undefined : <MountedContent inputRef={inputRef} onMount={onMount} />}
        </AutoSkeleton>
      </Card>
  </DemoCard>;
};

const AutoSkeletonPanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  return <>

    <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><AutoSkeletonExamples directionControl={<Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'direction' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="RTL direction" />} /></FluentProvider>
  </>;
};
export default AutoSkeletonPanel;
