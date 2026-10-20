import * as React from 'react';
import { makeStyles, mergeClasses } from '@griffel/react';
import { Skeleton, SkeletonItem } from '@fluentui/react-components';
import { AutoSkeleton } from '../../../packages/spfx-react-toolkit/src/controls/auto-skeleton/auto-skeleton';
import type { AutoSkeletonProps } from '../../../packages/spfx-react-toolkit/src/controls/auto-skeleton/auto-skeleton.types';
import { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
import * as width from '../../../packages/spfx-react-toolkit/src/helpers/styles/width';

interface AutoSkeletonFixtureConfig {
  loading: boolean;
  empty: boolean;
  narrow: boolean;
  animation: NonNullable<AutoSkeletonProps['animation']>;
  minSize: number;
  ignore: boolean;
  scrolling: boolean;
  secondLoading: boolean;
}
interface AutoSkeletonFixtureApi {
  render(patch: Partial<AutoSkeletonFixtureConfig>): void;
  focusAndLoad(): void;
  scrollRoot(top: number, left: number): void;
  readonly config: AutoSkeletonFixtureConfig;
}
declare global { interface Window { autoSkeletonFixture: AutoSkeletonFixtureApi; } }

const useStyles = makeStyles({
  root: { border: '3px solid currentColor', padding: '11px', boxSizing: 'border-box' },
  scroll: { height: '160px', overflow: 'auto' },
  content: { display: 'flex', flexDirection: 'column', gap: '12px' },
  rectangle: { width: '160px', height: '24px', borderRadius: '6px', backgroundColor: 'currentColor' },
  circle: { width: '40px', height: '40px', borderRadius: '50%', backgroundColor: 'currentColor' },
  tiny: { width: '4px', height: '4px', backgroundColor: 'currentColor' },
  hidden: { display: 'none' },
  visibleOverride: { visibility: 'visible' },
  text: { margin: 0 },
});
const ignored: readonly string[] = ['.as-selector-ignore'];
const noSelectors: readonly string[] = [];

/** Real geometry checks inherit the fixture's Fluent theme and writing direction. */
export function AutoSkeletonFixture(): React.ReactElement {
  const sx = useSx();
  const styles = useStyles();
  const [config, setConfig] = React.useState<AutoSkeletonFixtureConfig>({
    loading: false, empty: false, narrow: false, animation: 'wave', minSize: 8,
    ignore: true, scrolling: false, secondLoading: false,
  });
  const [value, setValue] = React.useState('persistent child value');
  const [clicks, setClicks] = React.useState(0);
  const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const inputRef = React.useMemo(() => React.createRef<HTMLInputElement>(), []);
  const update = React.useCallback((patch: Partial<AutoSkeletonFixtureConfig>) => setConfig(previous => ({ ...previous, ...patch })), []);
  React.useLayoutEffect(() => {
    window.autoSkeletonFixture = {
      render: update,
      get config() { return config; },
      focusAndLoad() { inputRef.current?.focus(); update({ loading: true }); },
      scrollRoot(top, left) {
        const root = rootRef.current;
        if (root) { root.scrollTop = top; root.scrollLeft = left; }
        window.dispatchEvent(new Event('resize'));
      },
    };
  }, [config, inputRef, rootRef, update]);
  return <section id="as-fixture" aria-label="AutoSkeleton browser checks">
    <h2>AutoSkeleton</h2>
    <button id="as-loading" type="button" onClick={() => update({ loading: !config.loading })}>Toggle loading</button>{' '}
    <button id="as-resize" type="button" onClick={() => update({ narrow: !config.narrow })}>Resize content</button>{' '}
    <button id="as-empty" type="button" onClick={() => update({ empty: !config.empty })}>Toggle empty</button>{' '}
    <button id="as-animation" type="button" onClick={() => update({ animation: config.animation === 'wave' ? 'pulse' : config.animation === 'pulse' ? 'none' : 'wave' })}>Toggle animation</button>{' '}
    <button id="as-focus-load" type="button" onClick={() => window.autoSkeletonFixture.focusAndLoad()}>Focus input then load</button>
    <button id="as-tab-before" type="button">Before content</button>
    <AutoSkeleton id="as-root" ref={rootRef} loading={config.loading} animation={config.animation} minSize={config.minSize}
      ignoreSelectors={config.ignore ? ignored : noSelectors} onClick={() => setClicks(previous => previous + 1)}
      className={mergeClasses(styles.root, config.scrolling && styles.scroll, sx(width.px(config.narrow ? 230 : 370)))}>
      {config.empty ? undefined : <div id="as-content-layout" data-skeleton-container="" className={styles.content}>
        <div id="as-rectangle" data-skeleton-shape="rectangle" className={styles.rectangle} />
        <div id="as-circle" data-skeleton-shape="circle" className={styles.circle} />
        <p id="as-text" className={styles.text}>Real wrapped text includes <strong id="as-inline">mixed inline content</strong> and more words to measure multiple lines without estimating their positions.</p>
        <input id="as-input" ref={inputRef} className={styles.visibleOverride} value={value} onChange={event => setValue(event.target.value)} />
        <div id="as-ignored" data-skeleton-ignore=""><div id="as-ignored-child" data-skeleton-shape="rectangle" className={styles.rectangle} /></div>
        <div id="as-selector-ignored" className="as-selector-ignore"><div id="as-selector-ignored-child" data-skeleton-shape="circle" className={styles.circle} /></div>
        <div id="as-hidden" className={styles.hidden}><div data-skeleton-shape="rectangle" className={styles.rectangle} /></div>
        <div id="as-tiny" className={styles.tiny} />
      </div>}
    </AutoSkeleton>
    <button id="as-tab-after" type="button">After content</button>
    <output id="as-clicks">{clicks}</output>
    <button id="as-ref" type="button" onClick={() => setClicks(rootRef.current?.tagName === 'DIV' ? 100 : -1)}>Inspect ref</button>
    <AutoSkeleton id="as-second" loading={config.secondLoading} className={sx(width.px(220))}>
      <div id="as-second-rectangle" data-skeleton-shape="rectangle" className={styles.rectangle} />
    </AutoSkeleton>
    <AutoSkeleton id="as-cold-empty" loading>{undefined}</AutoSkeleton>
    <Skeleton id="as-native" animation={config.animation === 'none' ? 'wave' : config.animation} aria-label="Native Fluent comparison">
      <SkeletonItem id="as-native-item" shape="rectangle" />
    </Skeleton>
  </section>;
}
