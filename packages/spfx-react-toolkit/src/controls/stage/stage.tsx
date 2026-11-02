/**
 * Explicit Dialog, Drawer and Popover surfaces sharing one caller-owned content structure.
 *
 * @packageDocumentation
 * @docPage controls/stage
 */
import * as React from 'react';
import { StageHeader } from './stage-header';
import { StageBody } from './stage-body';
import { StageFooter } from './stage-footer';
import { StageDialog } from './stage-dialog.internal';
import { StageDrawer } from './stage-drawer.internal';
import { StagePopover } from './stage-popover.internal';
import { useStageController } from './stage-controller.internal';
import { StageContext, useStageContextValue } from './stage-context.internal';
import type { StageController } from './stage-controller.internal';
import type { StageProps, StageSurfaceAttributes } from './stage.types';

/**
 * Keep the active anchor/placement when a different presentation is pending.
 * Common behavior remains live; the controller already owns state and compatible size.
 */
function useAdapterOptions(props: StageProps, controller: StageController): StageProps {
  const compatible = React.useRef(props);
  React.useLayoutEffect(() => {
    if (props.presentation === controller.presentation) compatible.current = props;
  }, [props, controller.presentation]);
  if (props.presentation === controller.presentation) return props;

  const previous = compatible.current;
  const state = props.open === undefined ? { defaultOpen: props.defaultOpen } : { open: props.open };
  const common = {
    ...state,
    children: props.children,
    onOpenChange: props.onOpenChange,
    dismissMode: props.dismissMode,
    unmountOnClose: props.unmountOnClose,
    mountNode: props.mountNode,
  };
  if (previous.presentation === 'popover') {
    return { ...common, presentation: 'popover', trigger: previous.trigger,
      positioning: previous.positioning, withArrow: previous.withArrow,
      size: controller.size === 'full' ? 'large' : controller.size };
  }
  if (previous.presentation === 'drawer') {
    return { ...common, presentation: 'drawer', trigger: previous.trigger,
      drawerPosition: previous.drawerPosition, size: controller.size };
  }
  return { ...common, presentation: 'dialog', trigger: previous.trigger, size: controller.size };
}

function getSurfaceAttributes(props: StageProps): StageSurfaceAttributes {
  const attributes: React.HTMLAttributes<HTMLDivElement> = { ...props };
  // Untyped consumers can still supply owned fields; never leak configuration,
  // raw HTML or a conflicting surface role/association into native markup.
  for (const key of ['presentation', 'size', 'trigger', 'drawerPosition', 'positioning', 'withArrow', 'open', 'defaultOpen',
    'onOpenChange', 'dismissMode', 'unmountOnClose', 'mountNode', 'children', 'title',
    'dangerouslySetInnerHTML', 'role', 'aria-modal', 'aria-labelledby', 'aria-describedby']) {
    Reflect.deleteProperty(attributes, key);
  }
  return attributes;
}

const StageRoot = React.forwardRef<HTMLDivElement, StageProps>(function StageRoot(props, ref) {
  const controller = useStageController(props);
  const options = useAdapterOptions(props, controller);
  const surfaceProps = getSurfaceAttributes(props);
  const context = useStageContextValue(controller, surfaceProps.id);
  const adapter = { controller, options, surfaceProps, surfaceRef: ref, children: props.children };
  return <StageContext.Provider value={context}>
    {controller.presentation === 'dialog' ? <StageDialog {...adapter} />
      : controller.presentation === 'drawer' ? <StageDrawer {...adapter} />
        : <StagePopover {...adapter} />}
  </StageContext.Provider>;
});
StageRoot.displayName = 'Stage';

/**
 * Present one composed content structure in a Dialog, Drawer or anchored Popover.
 * Use Header, Body and optional Footer in declared order; a caller-owned form wrapper
 * may contain all three. Presentation is explicit and latched for each open session.
 * Strings in Header stay literal with toolkit presets; caller JSX remains unchanged.
 * Requires shared Fluent theme/renderer, but no SPFx provider. Controlled consumers
 * accept or refuse requests by updating open; native event prevention is not a callback veto.
 * @param props - Surface attributes, presentation-specific options and open-state ownership.
 * @returns The integrated trigger, when supplied, and the Fluent-owned surface/portal.
 * @example Basic Stage composition
 * ```tsx
 * import * as React from 'react';
 * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * <Stage presentation="dialog" trigger={<button type="button">Edit</button>}>
 *   <Stage.Header title="Profile" closeLabel="Close" />
 *   <Stage.Body><input aria-label="Name" /></Stage.Body>
 *   <Stage.Footer><button type="button">Save</button></Stage.Footer>
 * </Stage>
 * ```
 * @useWhen One explicit modal or anchored presentation must host the same structured content.
 * @avoidWhen An open surface must switch immediately between presentation modes as the viewport changes; Stage latches presentation until the next session.
 * @throws Error - Presentation/size is invalid, full Popover is requested, or state ownership changes during a mount.
 * @example Profile editor with native Form ownership
 * ```tsx
 * import * as React from 'react';
 * import { createForm, Stage, useSx, alignItems, flex, flexItem, justifyContent, minHeight, width, gap } from '@apvee/spfx-react-toolkit';
 *
 * const Editor = createForm<{ name: string }>();
 *
 * export function ProfileEditor({ saveProfile }: { saveProfile: (values: { name: string }) => Promise<void> }): React.ReactElement {
 *   const [open, setOpen] = React.useState(false);
 *   const [discardRequested, setDiscardRequested] = React.useState(false);
 *   const methods = Editor.useForm({ defaultValues: { name: '' } });
 *   // Read during render to subscribe to RHF's proxied formState before callbacks use it.
 *   const { isDirty, isSubmitting } = methods.formState;
 *   const sx = useSx();
 *   return <Stage presentation="dialog" open={open}
 *     trigger={<button type="button">Edit profile</button>}
 *     onOpenChange={(_, data) => {
 *       if (!data.open && isSubmitting) return;
 *       if (!data.open && isDirty) { setDiscardRequested(true); return; }
 *       setOpen(data.open);
 *     }}>
 *     <Editor.Root methods={methods}
 *       className={sx(flex.column, flexItem.grow, minHeight.zero, width.full, gap.medium)}
 *       onSubmit={async values => {
 *         await saveProfile(values);
 *         methods.reset(values);
 *         setDiscardRequested(false);
 *         setOpen(false);
 *       }}>
 *       <Stage.Header title="Profile" closeLabel="Close" />
 *       <Stage.Body>
 *         <Editor.Input name="name" label="Display name" />
 *         {discardRequested && <div role="status">
 *           Discard changes?
 *           <button type="button" disabled={isSubmitting} onClick={() => {
 *             methods.reset(); setDiscardRequested(false); setOpen(false);
 *           }}>Discard</button>
 *           <button type="button" disabled={isSubmitting} onClick={() => setDiscardRequested(false)}>Keep editing</button>
 *         </div>}
 *       </Stage.Body>
 *       <Stage.Footer className={sx(flex.row, flex.wrap, alignItems.center, justifyContent.end, gap.small)}>
 *         <Editor.Submit appearance="primary">Save</Editor.Submit>
 *       </Stage.Footer>
 *     </Editor.Root>
 *   </Stage>;
 * }
 * ```
 * @remarks
 * ### Form ownership and pending work
 * Put the same-suite Form.Root inside the portaled surface so Body fields and Footer Submit share one native DOM form. React context from an outer Root alone cannot establish DOM form ownership. Submit through the suite's Root/Submit gate; direct methods.handleSubmit bypasses that gate. Read proxied isDirty/isSubmitting during rendering before callbacks use them. Await application persistence before resetting the baseline and closing. Keep close requests controlled while pending, and show inline discard confirmation rather than nesting a modal. Stage itself has no RHF dependency.
 * @example Edit a report with guarded Stage dismissal
 * ```tsx
 * import * as React from 'react';
 * import {
 *   Checkbox, Field, Select, Button,
 *   Input, makeStyles,
 * } from '@fluentui/react-components';
 * import { ArrowRight20Regular, Info24Regular } from '@fluentui/react-icons';
 * import {
 *   Stage, alignItems, createForm, flex, flexItem, gap, justifyContent, minHeight, minWidth, text, useSx, width,
 * } from '@apvee/spfx-react-toolkit';
 * import type {
 *   StageDismissMode, StageOpenChangeData, StageOpenChangeEvent, StagePresentation, StageProps, StageSize,
 * } from '@apvee/spfx-react-toolkit';
 *
 * interface Draft {
 *   name: string;
 *   category: string;
 * }
 * const Editor = createForm<Draft>();
 * const defaults: Draft = { name: 'Quarterly report', category: 'one' };
 * const useStyles = makeStyles({
 *   instance: { minInlineSize: '0' },
 *   form: { alignSelf: 'stretch', overflow: 'hidden' },
 *   output: { overflowWrap: 'anywhere' },
 *   narrowField: { minInlineSize: 0 },
 * });
 *
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   diagnostics?: React.ReactNode; children?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * const SurfaceNote: React.FC<{ revision: number }> = ({ revision }) => {
 *   const [value, setValue] = React.useState('Session note');
 *   const previous = React.useRef(revision);
 *   React.useEffect(() => { if (previous.current !== revision) { previous.current = revision; setValue('Session note'); } }, [revision]);
 *   return <Input aria-label="Surface note" input={{ as: 'input', ...{ 'data-stage-demo': 'surface-note' } }} value={value} onChange={(_, data) => setValue(data.value)} />;
 * };
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * }): React.ReactElement {
 *   const instance = 'A';
 *   const sx = useSx();
 *   const styles = useStyles();
 *   const titleSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const [slotObservation, setSlotObservation] = React.useState('not inspected');
 *   const controlWidth = sx(minWidth.zero, width.full);
 *   const narrowField = sx(styles.narrowField, minWidth.zero, width.full);
 *   const methods = Editor.useForm({ defaultValues: defaults });
 *   // Read during render to subscribe to RHF's formState proxy updates.
 *   const { isDirty, isSubmitting } = methods.formState;
 *   const [presentation, setPresentation] = React.useState<StagePresentation>('dialog');
 *   const [size, setSize] = React.useState<StageSize>('small');
 *   const [open, setOpen] = React.useState(false);
 *   const [longBody, setLongBody] = React.useState(false);
 *   const [longText, setLongText] = React.useState(false);
 *   const [title, setTitle] = React.useState('Report details');
 *   const [descriptionValue, setDescriptionValue] = React.useState('Update the document title and category.');
 *   const titleText = longText ? `Profile ${instance}: long literal title that must retain its accessible text while truncating visually. `.repeat(8) : title;
 *   const descriptionText = longText ? 'Long literal supporting description that keeps its full accessible content while using at most two visual lines. '.repeat(20) : descriptionValue;
 *   const [jsx, setJsx] = React.useState(false);
 *   const [icon, setIcon] = React.useState(true);
 *   const [description, setDescription] = React.useState(true);
 *   const [customStrings, setCustomStrings] = React.useState(false);
 *   const [closeButton, setCloseButton] = React.useState(true);
 *   const [footer, setFooter] = React.useState(true);
 *   const [retain, setRetain] = React.useState(false);
 *   const [dismissMode, setDismissMode] = React.useState<StageDismissMode>('standard');
 *   const [withArrow, setWithArrow] = React.useState(false);
 *   const [guardDirty, setGuardDirty] = React.useState(true);
 *   const [holdSave, setHoldSave] = React.useState(false);
 *   const [discard, setDiscard] = React.useState(false);
 *   const [request, setRequest] = React.useState('No request');
 *   const [status, setStatus] = React.useState('Not saved');
 *   const [saveCount, setSaveCount] = React.useState(0);
 *   const [openCount, setOpenCount] = React.useState(0);
 *   const [actionCount, setActionCount] = React.useState(0);
 *   const pendingSave = React.useRef<(() => void) | undefined>(undefined);
 *   const generation = React.useRef(0);
 *   const active = React.useRef(true);
 *   const previousReset = React.useRef(resetRevision);
 *   const [resetPending, setResetPending] = React.useState(false);
 *   const [contentRevision, setContentRevision] = React.useState(0);
 *   const requestReset = React.useCallback((): void => {
 *     generation.current += 1;
 *     setOpen(false);
 *     setResetPending(true);
 *     pendingSave.current?.();
 *     pendingSave.current = undefined;
 *   }, []);
 *   React.useEffect(() => () => {
 *     active.current = false;
 *     generation.current += 1;
 *     pendingSave.current?.();
 *     pendingSave.current = undefined;
 *   }, []);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     requestReset();
 *   }, [resetRevision, requestReset]);
 *   // Drain the real Form.Root submission gate before resetting RHF or permitting a new session.
 *   React.useEffect(() => {
 *     if (!resetPending || isSubmitting) return;
 *     methods.reset(defaults);
 *     setPresentation('dialog'); setSize('small'); setLongBody(false); setLongText(false);
 *     setTitle('Report details'); setDescriptionValue('Update the document title and category.');
 *     setJsx(false); setIcon(true); setDescription(true); setCustomStrings(false);
 *     setCloseButton(true); setFooter(true); setRetain(false); setDismissMode('standard');
 *     setWithArrow(false); setGuardDirty(true); setHoldSave(false); setDiscard(false);
 *     setRequest('No request'); setStatus('Not saved'); setSaveCount(0); setOpenCount(0); setActionCount(0);
 *     setSlotObservation('not inspected'); setContentRevision(value => value + 1); setResetPending(false);
 *   }, [resetPending, isSubmitting, methods]);
 *   React.useEffect(() => { if (open) setOpenCount(value => value + 1); }, [open]);
 *   const onOpenChange = (_event: StageOpenChangeEvent, data: StageOpenChangeData): void => {
 *     if (resetPending) return;
 *     setRequest(`${data.reason}: ${String(data.open)}`);
 *     if (!data.open && isSubmitting) {
 *       setStatus('Close refused: save pending');
 *       return;
 *     }
 *     if (!data.open && guardDirty && isDirty) {
 *       setDiscard(true);
 *       setStatus('Close refused: unsaved changes');
 *       return;
 *     }
 *     setDiscard(false);
 *     setOpen(data.open);
 *   };
 *   const requestClose = (event: React.MouseEvent<HTMLElement>): void => {
 *     onOpenChange(event, { open: false, reason: 'closeButton' });
 *   };
 *   // Capture the submission epoch before RHF can defer invoking this callback.
 *   const submissionGeneration = generation.current;
 *   const save = async (value: Draft): Promise<void> => {
 *     const currentGeneration = submissionGeneration;
 *     if (!active.current || currentGeneration !== generation.current) return;
 *     setSaveCount(count => count + 1);
 *     setStatus('Saving');
 *     if (holdSave) {
 *       await new Promise<void>(resolve => { pendingSave.current = resolve; });
 *       if (currentGeneration === generation.current) pendingSave.current = undefined;
 *     }
 *     if (!active.current || currentGeneration !== generation.current) return;
 *     methods.reset(value);
 *     setDiscard(false);
 *     setStatus(`Saved ${value.name}`);
 *   };
 *   const changePresentation = (value: string): void => {
 *     if (value !== 'dialog' && value !== 'drawer' && value !== 'popover') return;
 *     if (value === 'popover' && size === 'full') setSize('small');
 *     setPresentation(value);
 *   };
 *   const trigger = <Button type="button" data-stage-demo="trigger" disabled={resetPending || isSubmitting} appearance="primary">Open editor</Button>;
 *   const surfaceOptions: Extract<StageProps, { open: boolean }> = presentation === 'popover'
 *     ? { presentation, size: size === 'full' ? 'small' : size, withArrow, trigger, open, children: undefined }
 *     : { presentation, size, trigger, open, children: undefined };
 *
 *   return <section data-stage-demo="panel"><section data-stage-demo="instance" data-stage-instance={instance} className={sx(styles.instance, flex.column, gap.medium)}>
 *     {render({ controls: <>
 *       <Field label="Presentation"><Select data-demo-variant="Presentation" data-stage-demo="mode" value={presentation} disabled={open} onChange={event => changePresentation(event.target.value)}>
 *         <option value="dialog">Dialog</option><option value="drawer">Drawer</option><option value="popover">Popover</option>
 *       </Select></Field>
 *       <Field label="Size"><Select data-stage-demo="size" value={size} onChange={event => {
 *         const value = event.target.value;
 *         if (value === 'small' || value === 'medium' || value === 'large' || (value === 'full' && presentation !== 'popover')) setSize(value);
 *       }}>
 *         <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *         {presentation !== 'popover' && <option value="full">Full</option>}
 *       </Select></Field>
 *       <Field label="Title"><Input input={{ as: 'input', ...{ 'data-stage-demo': 'title' } }} value={title} onChange={(_, data) => setTitle(data.value)} /></Field>
 *       <Field label="Description"><Input input={{ as: 'input', ...{ 'data-stage-demo': 'description-input' } }} value={descriptionValue} onChange={(_, data) => setDescriptionValue(data.value)} /></Field>
 *     </>, commands: undefined, status: <>
 *     <output data-stage-demo="status">{status}</output>
 *     <output data-stage-demo="open">Open: {String(open)}</output>
 *     </>, diagnostics: <>
 *       <Field label="Dismiss"><Select data-stage-demo="dismiss" value={dismissMode} onChange={event => setDismissMode(event.target.value === 'explicit' ? 'explicit' : 'standard')}>
 *         <option value="standard">Standard</option><option value="explicit">Explicit</option>
 *       </Select></Field>
 *       <Checkbox data-stage-demo="long-body" label="Long body" checked={longBody} onChange={(_, data) => setLongBody(data.checked === true)} />
 *       <Checkbox data-stage-demo="jsx" label="JSX title/description" checked={jsx} onChange={(_, data) => setJsx(data.checked === true)} />
 *       <Checkbox data-stage-demo="long-text" label="Long title and description" checked={longText} onChange={(_, data) => setLongText(data.checked === true)} />
 *       <Checkbox data-stage-demo="icon" label="Icon" checked={icon} onChange={(_, data) => setIcon(data.checked === true)} />
 *       <Checkbox data-stage-demo="description" label="Description" checked={description} onChange={(_, data) => setDescription(data.checked === true)} />
 *       <Checkbox data-stage-demo="strings" label="Custom close text" checked={customStrings} onChange={(_, data) => setCustomStrings(data.checked === true)} />
 *       <Checkbox data-stage-demo="close" label="Header X" checked={closeButton} onChange={(_, data) => setCloseButton(data.checked === true)} />
 *       <Checkbox data-stage-demo="footer-toggle" label="Footer" checked={footer} onChange={(_, data) => setFooter(data.checked === true)} />
 *       <Checkbox data-stage-demo="retention" label="Retain after close" checked={retain} onChange={(_, data) => setRetain(data.checked === true)} />
 *       <Checkbox data-stage-demo="arrow-toggle" label="Popover arrow" checked={withArrow} disabled={presentation !== 'popover'} onChange={(_, data) => setWithArrow(data.checked === true)} />
 *       <Checkbox data-stage-demo="dirty-guard" label="Guard unsaved changes" checked={guardDirty} onChange={(_, data) => setGuardDirty(data.checked === true)} />
 *       <Checkbox data-stage-demo="hold-save" label="Complete save manually" checked={holdSave} onChange={(_, data) => setHoldSave(data.checked === true)} />
 *       <Button data-stage-demo="inspect-slot" onClick={() => setSlotObservation(titleSlotRef.current?.tagName ?? 'absent')}>Inspect title</Button><output data-stage-demo="slot-observation">{slotObservation}</output>
 *
 *     <output data-stage-demo="request" className={styles.output}>{request}</output>
 *     <output data-stage-demo="save-count">Save callbacks: {saveCount}</output>
 *     <output data-stage-demo="mount-count">Open sessions: {openCount}</output>
 *     </>, children: <Stage {...surfaceOptions} open={open} onOpenChange={onOpenChange} dismissMode={dismissMode} unmountOnClose={!retain}
 *       data-stage-demo="surface" data-stage-instance={instance}>
 *       <Editor.Root methods={methods} onSubmit={save} data-stage-demo="form"
 *         className={sx(styles.form, flex.column, flexItem.grow, minHeight.zero, width.full, gap.medium)}>
 *         <Stage.Header data-stage-demo="header" title={{ ref: titleSlotRef, 'data-stage-demo': 'title-slot', children: jsx ? <em data-stage-demo="jsx-title">{titleText}</em> : titleText }}
 *           description={description ? (jsx ? <span data-stage-demo="jsx-description">Editable <strong>profile</strong></span> : descriptionText) : undefined}
 *           icon={icon ? { children: (Tag, props) => <Tag {...props}><Info24Regular data-stage-demo="callback-icon" width={80} height={80} /></Tag> } : undefined} showCloseButton={closeButton} strings={customStrings ? { closeLabel: `Dismiss ${instance}` } : undefined} closeLabel={customStrings ? undefined : `Close ${instance}`}
 *           action={{ children: (Tag, props) => <Tag {...props} data-stage-demo="action-slot">
 *             <Button type="button" className={sx(minWidth.zero)} icon={<ArrowRight20Regular />} data-stage-demo="header-action" onClick={() => setActionCount(value => value + 1)}>
 *               <span className={sx(minWidth.zero, text.truncate)}>Share</span>
 *             </Button>
 *           </Tag> }} />
 *         <Stage.Body data-stage-demo="body">
 *           <div className={sx(flex.column, gap.medium)}>
 *             <Editor.Input name="name" label="Document title" data-stage-demo="name"
 *               className={narrowField} root={{ className: controlWidth }} input={{ className: sx(minWidth.zero) }} />
 *             <div data-stage-demo="combobox"><Editor.Combobox name="category" label="Category"
 *               className={narrowField} root={{ className: controlWidth }} input={{ className: sx(minWidth.zero) }}
 *               options={[{ value: 'one', label: 'Report' }, { value: 'two', label: 'Project' }]} /></div>
 *             <SurfaceNote revision={contentRevision} />
 *             <Button type="button" data-stage-demo="reset-editor" onClick={requestReset}>Reset editor</Button>
 *             <output data-stage-demo="dirty">Dirty: {String(isDirty)}</output>
 *             <output data-stage-demo="pending">Pending: {String(isSubmitting)}</output>
 *             {isSubmitting && <Button type="button" data-stage-demo="complete-save" onClick={() => pendingSave.current?.()}>Complete save</Button>}
 *             <output data-stage-demo="action-count">Header actions: {actionCount}</output>
 *             {discard && <div role="status" data-stage-demo="discard-prompt">
 *               <p>Discard unsaved changes?</p>
 *               <Button type="button" data-stage-demo="discard" disabled={isSubmitting} onClick={() => {
 *                 methods.reset(); setDiscard(false); setOpen(false); setStatus('Discarded changes');
 *               }}>Discard and close</Button>
 *               <Button type="button" data-stage-demo="keep-editing" onClick={() => setDiscard(false)}>Keep editing</Button>
 *             </div>}
 *             {!footer && <Button type="button" data-stage-demo="body-close" onClick={requestClose}>Close editor</Button>}
 *             {longBody && Array.from({ length: 35 }, (_, index) => <p key={index}>Scrollable profile note {index + 1}. Header and footer remain available while this body scrolls.</p>)}
 *           </div>
 *         </Stage.Body>
 *         {footer && <Stage.Footer data-stage-demo="footer" className={sx(flex.row, flex.wrap, alignItems.center, justifyContent.end, gap.small)}>
 *           <Editor.Submit data-stage-demo="save" appearance="primary">Save</Editor.Submit>
 *           <Button type="button" data-stage-demo="footer-close" onClick={requestClose}>Close editor</Button>
 *         </Stage.Footer>}
 *       </Editor.Root>
 *     </Stage> })}
 *   </section></section>;
 * }
 *
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Edit the report in each presentation; dirty or pending saves refuse dismissal, and accepted closes restore the native trigger focus.',
 * };
 * ```
 */
export const Stage = Object.assign(StageRoot, {
  /**
   * Semantic heading, decorative icon, supporting text and caller-owned action. Requires the owning Stage.
   * @param props - Native div attributes and header content.
   * @returns The header div within the active surface; its ref targets that div.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
   *
   * <Stage presentation="dialog"><Stage.Header title="Profile" /></Stage>
   * ```
   */
  Header: StageHeader,
  /**
   * Growing scroll region, preserving the consumer's content. Requires the owning Stage.
   * @param props - Native div attributes and body content.
   * @returns The body div within the active surface; its ref targets that div.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
   *
   * <Stage presentation="dialog"><Stage.Body>Content</Stage.Body></Stage>
   * ```
   */
  Body: StageBody,
  /**
   * Optional free-content footer; actions do not dismiss automatically. Requires the owning Stage.
   * @param props - Native div attributes and footer content.
   * @returns The footer div within the active surface; its ref targets that div.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
   *
   * <Stage presentation="dialog"><Stage.Footer><button type="button">Save</button></Stage.Footer></Stage>
   * ```
   */
  Footer: StageFooter,
});
