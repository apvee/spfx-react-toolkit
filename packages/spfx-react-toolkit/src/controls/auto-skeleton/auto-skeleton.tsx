/**
 * Geometry-derived loading placeholders over mounted, bounded HTML content.
 *
 * @packageDocumentation
 * @docPage controls/auto-skeleton
 */
import * as React from 'react';
import { mergeClasses, Skeleton, SkeletonItem } from '@fluentui/react-components';
import { useIsomorphicLayoutEffect, useMergedRefs } from '@fluentui/react-utilities';
import { useAutoSkeletonStyles, useAutoSkeletonGeometryClasses } from './auto-skeleton.styles.internal';
import { useAutoSkeleton } from './use-auto-skeleton.internal';
import type { AutoSkeletonProps } from './auto-skeleton.types';
import type { AutoSkeletonBone } from './auto-skeleton-geometry.internal';

const emptySelectors: readonly string[] = [];
const emptyGeometry: readonly AutoSkeletonBone[] = [];
/**
 * Derive Fluent skeleton items from the live geometry of mounted HTML content.
 * Children stay mounted in a block wrapper. Real text rectangles and visual leaves
 * produce only circle/rectangle items; text is inset by up to 2px per side for spacing.
 * Rectangles retain Fluent's rounded corners when content has no measured radius;
 * positive content radii are preserved and circles use a 50% radius.
 * Animation defaults to Fluent wave; `animation="none"` uses static themed Fluent items
 * without shimmer, while preserving loading visibility, focus and aria-busy behavior.
 * No geometry cache or generated inline styles.
 * Loading content is hidden/inert only when measurable and native inert is supported.
 * Otherwise content remains available with aria-busy. Native inert and DOM geometry
 * are browser requirements; server imports are safe but server rendering has no overlay.
 * Bounded HTML only: portals, shadow DOM, transforms and nested clipping are unsupported.
 * @param props - Loading state, scan configuration and native div attributes.
 * @returns A div root containing mounted content and an optional decorative skeleton.
 * @example Loading a mounted profile
 * ```tsx
 * import * as React from 'react';
 * import { AutoSkeleton } from '@apvee/spfx-react-toolkit';
 * const isLoading = true;
 * const photoUrl = '/images/profile.png';
 * const name = 'Ada';
 * <AutoSkeleton loading={isLoading} animation="pulse">
 *   <section><img data-skeleton-shape="circle" src={photoUrl} alt="" /><p>{name}</p></section>
 * </AutoSkeleton>
 * ```
 *
 * @example Static placeholders
 * ```tsx
 * import * as React from 'react';
 * import { AutoSkeleton } from '@apvee/spfx-react-toolkit';
 * const showPlaceholder = true;
 * function ProfileCard(): React.ReactElement { return <article>Ada — project owner</article>; }
 * <AutoSkeleton loading={showPlaceholder} animation="none"><ProfileCard /></AutoSkeleton>
 * ```
 *
 * @useWhen Representative mounted HTML exists and a loading overlay should follow its measured geometry.
 * @avoidWhen Use explicit skeleton placeholders for cold empty data, virtualized content or unsupported transformed geometry.
 *
 *
 * @example Compose shape container and conditional ignore markers
 * ```tsx
 * import * as React from 'react';
 * import { Card, CardHeader, Text } from '@fluentui/react-components';
 * import { AutoSkeleton, useSx, flex, gap } from '@apvee/spfx-react-toolkit';
 * function ProfileLoading({ loading, showSummary, ignoreNotes }: { loading: boolean; showSummary: boolean; ignoreNotes: boolean }) {
 *   const sx = useSx();
 *   return <div>
 *     <AutoSkeleton loading={loading}>
 *       <Card data-skeleton-shape="rectangle"><CardHeader header={<Text>Profile</Text>} /><Text>Profile details</Text></Card>
 *     </AutoSkeleton>
 *     <AutoSkeleton loading={loading}>
 *       <div data-skeleton-container="" className={sx(flex.column, gap.medium)}>
 *         {showSummary && <p>Profile summary</p>}
 *         <p data-skeleton-ignore={ignoreNotes ? '' : undefined}>Notes excluded when requested</p>
 *       </div>
 *     </AutoSkeleton>
 *   </div>;
 * }
 * ```
 *
 * @example Explore mounted profile loading and native focus
 * ```tsx
 * // This standalone React scenario uses the host Fluent theme and Griffel renderer. The mount counter measures real content mounts; Reset restores the note without remounting retained content.
 * import * as React from 'react';
 * import { Avatar, Card, Text, FluentProvider, makeStyles, mergeClasses, tokens, Checkbox, Select, Button, Field, Input } from '@fluentui/react-components';
 * import { Calendar16Regular, Person16Regular } from '@fluentui/react-icons';
 * import { AutoSkeleton, AutoSkeletonProps, useSx, useStableCallback, alignItems, flex, gap, width, background, minWidth } from '@apvee/spfx-react-toolkit';
 *
 * const useStyles = makeStyles({
 *   controls: { display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: tokens.spacingVerticalM, minWidth: 0,
 *     '& > *': { minWidth: 0, maxWidth: '100%' }, '@media (max-width: 600px)': { gridTemplateColumns: 'minmax(0, 1fr)' } },
 *   controlRow: { gridColumn: '1 / -1' },
 *
 *   preview: { boxSizing: 'border-box', maxWidth: '100%' },
 *   rectangle: { width: '160px', height: '24px', borderRadius: '4px', maxWidth: '100%' },
 *   circle: { width: '40px', height: '40px', flexShrink: 0 },
 *   secondary: { color: tokens.colorNeutralForeground2 },
 *   paragraph: { margin: 0 },
 *   visibleOverride: { visibility: 'visible' },
 * });
 * const ignoredSelectors: readonly string[] = ['.auto-skeleton-demo-ignore'];
 * const noSelectors: readonly string[] = [];
 * const animations: readonly NonNullable<AutoSkeletonProps['animation']>[] = ['wave', 'pulse', 'none'];
 *
 * const MountedContent: React.FC<{ inputRef: React.RefObject<HTMLInputElement>; onMount: () => void; resetRevision: number }> = ({ inputRef, onMount, resetRevision }) => {
 *   const sx = useSx();
 *   const styles = useStyles();
 *   const [note, setNote] = React.useState('Documentation update');
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision; setNote('Documentation update');
 *   }, [resetRevision]);
 *   React.useEffect(() => { onMount(); }, [onMount]);
 *   return <div data-skeleton-container="" className={sx(flex.column, gap.medium)}>
 *     <div data-skeleton-container="" className={sx(flex.row, alignItems.center, gap.medium, minWidth.zero)}>
 *       <Avatar name="Giulia Rossi" size={40} data-auto-skeleton-demo="circle" data-skeleton-shape="circle" className={styles.circle} />
 *       <div data-skeleton-container="" className={sx(flex.column, minWidth.zero)}>
 *         <Text block weight="semibold" size={400} data-auto-skeleton-demo="rectangle" data-skeleton-shape="rectangle" className={styles.rectangle}>Giulia Rossi</Text>
 *         <Text size={200} className={styles.secondary}>Project coordinator</Text>
 *       </div>
 *     </div>
 *     <Text as="p" data-auto-skeleton-demo="mixed-text" className={styles.paragraph}>Giulia Rossi coordinates the <strong>intranet project</strong>. The team is completing the documentation for the next review.</Text>
 *     <Field label="Project note"><Input input={{ as: 'input', ref: inputRef, ...{ 'data-auto-skeleton-demo': 'focus-target' } }}  className={styles.visibleOverride}
 *       value={note} onChange={(_event, data) => setNote(data.value)} /></Field>
 *     <div data-skeleton-ignore="" className={sx(flex.row, alignItems.center, gap.small)}><Person16Regular aria-hidden="true" /><Text size={200}>Team lead: Giulia Rossi.</Text></div>
 *     <div className={mergeClasses('auto-skeleton-demo-ignore', sx(flex.row, alignItems.center, gap.small))}><Calendar16Regular aria-hidden="true" /><Text size={200}>Next review: Friday.</Text></div>
 *   </div>;
 * };
 *
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   diagnostics?: React.ReactNode; children?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * }): React.ReactElement {
 *   const [rtl, setRtl] = React.useState(false);
 *   const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const directionControl = <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'direction' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="RTL direction" />;
 *   const sx = useSx();
 *   const demoStyles = useStyles();
 *   const styles = useStyles();
 *   const [loading, setLoading] = React.useState(false);
 *   const [empty, setEmpty] = React.useState(false);
 *   const [narrow, setNarrow] = React.useState(false);
 *   const [ignore, setIgnore] = React.useState(true);
 *   const [animation, setAnimation] = React.useState<NonNullable<AutoSkeletonProps['animation']>>('wave');
 *   const [minSize, setMinSize] = React.useState(8);
 *   const [mounts, setMounts] = React.useState(0);
 *   const [clicks, setClicks] = React.useState(0);
 *   const [observed, setObserved] = React.useState('not inspected');
 *   const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const inputRef = React.useMemo(() => React.createRef<HTMLInputElement>(), []);
 *   const onMount = useStableCallback(() => setMounts(value => value + 1));
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setLoading(false); setEmpty(false); setNarrow(false); setIgnore(true); setAnimation('wave'); setMinSize(8);
 *     setClicks(0); setObserved('not inspected'); setRtl(false);
 *   }, [resetRevision]);
 *   return render({ controls: <><Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'loading' } }} checked={loading} onChange={(_event, data) => setLoading(data.checked === true)} label="Loading" />
 * <Field label="Animation" hint="Choose none for static placeholders. Loading still controls visibility and busy state."><Select select={{ as: 'select', ...{ 'data-auto-skeleton-demo': 'animation', 'data-demo-variant': 'auto-skeleton' } }} value={animation} onChange={(_event, data) => {
 *         const selected = animations.find(value => value === data.value); if (selected) setAnimation(selected);
 *       }}>{animations.map(value => <option key={value}>{value}</option>)}</Select></Field>
 * <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'resize' } }} checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} label="Narrow width" />
 * <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'empty' } }} checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} label="Empty content" /></>,
 *     commands: <>
 *       <Button data-auto-skeleton-demo="focus-load" type="button" disabled={loading || empty} onClick={() => {
 *         inputRef.current?.focus(); setLoading(true);
 *       }}>Focus the note and load</Button>
 *     </>,
 *     status: <p aria-live="polite">{loading ? 'Loading' : 'Loaded'} · Content mounts: {mounts} · Preview clicks: {clicks}</p>,
 *     diagnostics: <><div className={demoStyles.controls}>{directionControl}
 * <Field label="Minimum size"><Select select={{ as: 'select', ...{ 'data-auto-skeleton-demo': 'minSize' } }} value={minSize} onChange={(_event, data) => setMinSize(Number(data.value))}>
 *         {[8, 16, 32].map(value => <option key={value} value={value}>{value}px</option>)}
 *       </Select></Field>
 * <Checkbox input={{ as: 'input', ...{ 'data-auto-skeleton-demo': 'ignoreSelectors' } }} checked={ignore} onChange={(_event, data) => setIgnore(data.checked === true)} label="Ignore selected content" /></div>
 * <Button data-auto-skeleton-demo="inspect-ref" type="button" onClick={() => {
 *         const root = rootRef.current;
 *         setObserved(root ? `${root.tagName.toLowerCase()}, ${root.querySelectorAll('[data-auto-skeleton-shape]').length} items, busy=${root.getAttribute('aria-busy')}` : 'unmounted');
 *       }}>Inspect root ref</Button>
 *
 * <p data-auto-skeleton-demo="observations">Root: {observed}. Content mounts: {mounts}. Preview clicks: {clicks}.</p>
 *
 * <p>Type a note, toggle loading, and return to loaded content: the value and mount count stay unchanged. Empty content has no inferred placeholder. Native inert support is required to hide content safely.</p>
 * </>,
 *     children: <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><Card data-auto-skeleton-demo="card" appearance="filled" size="medium" focusMode="off" className={mergeClasses(styles.preview, sx(width.px(narrow ? 220 : 360)))}>
 *         <AutoSkeleton data-auto-skeleton-demo="preview" ref={rootRef} loading={loading} animation={animation} minSize={minSize}
 *           ignoreSelectors={ignore ? ignoredSelectors : noSelectors} className={sx(width.full, minWidth.zero)}
 *           onClick={() => setClicks(value => value + 1)}>
 *           {empty ? undefined : <MountedContent inputRef={inputRef} onMount={onMount} resetRevision={resetRevision} />}
 *         </AutoSkeleton>
 *       </Card></FluentProvider> });
 * }
 *
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Type a note, vary loading and geometry, and observe retained content, native focus and measured shapes.' };
 * ```
 */
export const AutoSkeleton = React.forwardRef<HTMLDivElement, AutoSkeletonProps>((props, forwardedRef) => {
  const { loading, children, animation = 'wave', minSize = 8, ignoreSelectors = emptySelectors, className, ...nativeProps } = props;
  delete nativeProps.dangerouslySetInnerHTML;
  const [root, setRoot] = React.useState<HTMLDivElement>();
  const [content, setContent] = React.useState<HTMLDivElement>();
  const rootCallback = React.useCallback((element: HTMLDivElement | null) => setRoot(element ?? undefined), []);
  const contentCallback = React.useCallback((element: HTMLDivElement | null) => setContent(element ?? undefined), []);
  const rootRef = useMergedRefs(rootCallback, forwardedRef);
  const styles = useAutoSkeletonStyles();
  const isStatic = animation === 'none';
  const bones = useAutoSkeleton({ root, content, loading, children, minSize, ignoreSelectors });
  const active = loading && bones.length > 0 && content !== undefined && 'inert' in content;
  const geometryClasses = useAutoSkeletonGeometryClasses(active ? bones : emptyGeometry);
  useIsomorphicLayoutEffect(() => {
    if (!active || !content) return;
    const focused = content.ownerDocument.activeElement;
    if (focused && content.contains(focused) && 'blur' in focused && typeof focused.blur === 'function') focused.blur();
  }, [active, content]);
  return <div {...nativeProps} ref={rootRef} aria-busy={loading} className={mergeClasses(styles.root, className)}>
    <div ref={contentCallback} data-auto-skeleton-content="" className={active ? styles.contentHidden : undefined}
      aria-hidden={active ? true : undefined} {...{ inert: active ? '' : undefined }}>{children}</div>
    {active && <Skeleton data-auto-skeleton-overlay="" animation={isStatic ? 'wave' : animation} aria-hidden="true" aria-busy={undefined}
      role="presentation" className={styles.overlay}>
      {bones.map((bone, index) => <SkeletonItem key={`${bone.x}:${bone.y}:${bone.width}:${bone.height}:${bone.radius}:${bone.shape}`}
        data-auto-skeleton-shape={bone.shape} shape={bone.shape} className={mergeClasses(styles.item, isStatic && styles.staticItem, geometryClasses[index])} />)}
    </Skeleton>}
  </div>;
});
AutoSkeleton.displayName = 'AutoSkeleton';
