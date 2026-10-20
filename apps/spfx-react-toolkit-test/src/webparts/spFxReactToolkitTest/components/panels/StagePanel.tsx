import * as React from 'react';
import { DemoCard } from '../shared';
import {
  Checkbox, Field, Select, Button,
  Input, makeStyles,
} from '@fluentui/react-components';
import { ArrowRight20Regular, Info24Regular } from '@fluentui/react-icons';
import {
  Stage, alignItems, createForm, flex, flexItem, gap, justifyContent, minHeight, minWidth, text, useSx, width,
} from '@apvee/spfx-react-toolkit';
import type {
  StageDismissMode, StageOpenChangeData, StageOpenChangeEvent, StagePresentation, StageProps, StageSize,
} from '@apvee/spfx-react-toolkit';

interface Draft {
  name: string;
  category: string;
}
const Editor = createForm<Draft>();
const defaults: Draft = { name: 'Quarterly report', category: 'one' };
const useStyles = makeStyles({
  instance: { minInlineSize: '0' },
  form: { alignSelf: 'stretch', overflow: 'hidden' },
  output: { overflowWrap: 'anywhere' },
  narrowField: { minInlineSize: 0 },
});

export const StageExample: React.FC<{ instance: string }> = ({ instance }) => {
  const sx = useSx();
  const styles = useStyles();
  const titleSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const [slotObservation, setSlotObservation] = React.useState('not inspected');
  const controlWidth = sx(minWidth.zero, width.full);
  const narrowField = sx(styles.narrowField, minWidth.zero, width.full);
  const methods = Editor.useForm({ defaultValues: defaults });
  // Read during render to subscribe to RHF's formState proxy updates.
  const { isDirty, isSubmitting } = methods.formState;
  const [presentation, setPresentation] = React.useState<StagePresentation>('dialog');
  const [size, setSize] = React.useState<StageSize>('small');
  const [open, setOpen] = React.useState(false);
  const [longBody, setLongBody] = React.useState(false);
  const [longText, setLongText] = React.useState(false);
  const [title, setTitle] = React.useState('Report details');
  const [descriptionValue, setDescriptionValue] = React.useState('Update the document title and category.');
  const titleText = longText ? `Profile ${instance}: long literal title that must retain its accessible text while truncating visually. `.repeat(8) : title;
  const descriptionText = longText ? 'Long literal supporting description that keeps its full accessible content while using at most two visual lines. '.repeat(20) : descriptionValue;
  const [jsx, setJsx] = React.useState(false);
  const [icon, setIcon] = React.useState(true);
  const [description, setDescription] = React.useState(true);
  const [closeButton, setCloseButton] = React.useState(true);
  const [footer, setFooter] = React.useState(true);
  const [retain, setRetain] = React.useState(false);
  const [dismissMode, setDismissMode] = React.useState<StageDismissMode>('standard');
  const [withArrow, setWithArrow] = React.useState(false);
  const [guardDirty, setGuardDirty] = React.useState(true);
  const [holdSave, setHoldSave] = React.useState(false);
  const [discard, setDiscard] = React.useState(false);
  const [request, setRequest] = React.useState('No request');
  const [status, setStatus] = React.useState('Not saved');
  const [saveCount, setSaveCount] = React.useState(0);
  const [mountCount, setMountCount] = React.useState(0);
  const [actionCount, setActionCount] = React.useState(0);
  const pendingSave = React.useRef<(() => void) | undefined>(undefined);
  React.useEffect(() => { if (open) setMountCount(value => value + 1); }, [open]);
  const onOpenChange = (_event: StageOpenChangeEvent, data: StageOpenChangeData): void => {
    setRequest(`${data.reason}: ${String(data.open)}`);
    if (!data.open && isSubmitting) {
      setStatus('Close refused: save pending');
      return;
    }
    if (!data.open && guardDirty && isDirty) {
      setDiscard(true);
      setStatus('Close refused: unsaved changes');
      return;
    }
    setDiscard(false);
    setOpen(data.open);
  };
  const requestClose = (event: React.MouseEvent<HTMLElement>): void => {
    onOpenChange(event, { open: false, reason: 'closeButton' });
  };
  const save = async (value: Draft): Promise<void> => {
    setSaveCount(count => count + 1);
    setStatus('Saving');
    if (holdSave) {
      await new Promise<void>(resolve => { pendingSave.current = resolve; });
      pendingSave.current = undefined;
    }
    methods.reset(value);
    setDiscard(false);
    setStatus(`Saved ${value.name}`);
  };
  const changePresentation = (value: string): void => {
    if (value !== 'dialog' && value !== 'drawer' && value !== 'popover') return;
    if (value === 'popover' && size === 'full') setSize('small');
    setPresentation(value);
  };
  const trigger = <Button type="button" data-stage-demo="trigger" appearance="primary">Open editor</Button>;
  const surfaceOptions: Extract<StageProps, { open: boolean }> = presentation === 'popover'
    ? { presentation, size: size === 'full' ? 'small' : size, withArrow, trigger, open, children: undefined }
    : { presentation, size, trigger, open, children: undefined };

  return <section data-stage-demo="instance" data-stage-instance={instance} className={sx(styles.instance, flex.column, gap.medium)}>
    <DemoCard title="Stage" iconName="Edit" description="Edit the report in a dialog, drawer or popover." controls={<>
      <Field label="Presentation"><Select data-demo-variant="Presentation" data-stage-demo="mode" value={presentation} disabled={open} onChange={event => changePresentation(event.target.value)}>
        <option value="dialog">Dialog</option><option value="drawer">Drawer</option><option value="popover">Popover</option>
      </Select></Field>
      <Field label="Size"><Select data-stage-demo="size" value={size} onChange={event => {
        const value = event.target.value;
        if (value === 'small' || value === 'medium' || value === 'large' || (value === 'full' && presentation !== 'popover')) setSize(value);
      }}>
        <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
        {presentation !== 'popover' && <option value="full">Full</option>}
      </Select></Field>
      <Field label="Title"><Input input={{ as: 'input', ...{ 'data-stage-demo': 'title' } }} value={title} onChange={(_, data) => setTitle(data.value)} /></Field>
      <Field label="Description"><Input input={{ as: 'input', ...{ 'data-stage-demo': 'description-input' } }} value={descriptionValue} onChange={(_, data) => setDescriptionValue(data.value)} /></Field>
    </>} commands={undefined} status={<>
    <output data-stage-demo="status">{status}</output>
    <output data-stage-demo="open">Open: {String(open)}</output>
    </>} diagnostics={<>
      <Field label="Dismiss"><Select data-stage-demo="dismiss" value={dismissMode} onChange={event => setDismissMode(event.target.value === 'explicit' ? 'explicit' : 'standard')}>
        <option value="standard">Standard</option><option value="explicit">Explicit</option>
      </Select></Field>
      <Checkbox data-stage-demo="long-body" label="Long body" checked={longBody} onChange={(_, data) => setLongBody(data.checked === true)} />
      <Checkbox data-stage-demo="jsx" label="JSX title/description" checked={jsx} onChange={(_, data) => setJsx(data.checked === true)} />
      <Checkbox data-stage-demo="long-text" label="Long title and description" checked={longText} onChange={(_, data) => setLongText(data.checked === true)} />
      <Checkbox data-stage-demo="icon" label="Icon" checked={icon} onChange={(_, data) => setIcon(data.checked === true)} />
      <Checkbox data-stage-demo="description" label="Description" checked={description} onChange={(_, data) => setDescription(data.checked === true)} />
      <Checkbox data-stage-demo="close" label="Header X" checked={closeButton} onChange={(_, data) => setCloseButton(data.checked === true)} />
      <Checkbox data-stage-demo="footer-toggle" label="Footer" checked={footer} onChange={(_, data) => setFooter(data.checked === true)} />
      <Checkbox data-stage-demo="retention" label="Retain after close" checked={retain} onChange={(_, data) => setRetain(data.checked === true)} />
      <Checkbox data-stage-demo="arrow-toggle" label="Popover arrow" checked={withArrow} disabled={presentation !== 'popover'} onChange={(_, data) => setWithArrow(data.checked === true)} />
      <Checkbox data-stage-demo="dirty-guard" label="Guard unsaved changes" checked={guardDirty} onChange={(_, data) => setGuardDirty(data.checked === true)} />
      <Checkbox data-stage-demo="hold-save" label="Complete save manually" checked={holdSave} onChange={(_, data) => setHoldSave(data.checked === true)} />
      <Button data-stage-demo="inspect-slot" onClick={() => setSlotObservation(titleSlotRef.current?.tagName ?? 'absent')}>Inspect title</Button><output data-stage-demo="slot-observation">{slotObservation}</output>

    <output data-stage-demo="request" className={styles.output}>{request}</output>
    <output data-stage-demo="save-count">Save callbacks: {saveCount}</output>
    <output data-stage-demo="mount-count">Surface mounts: {mountCount}</output>
    </>}>
    <Stage {...surfaceOptions} open={open} onOpenChange={onOpenChange} dismissMode={dismissMode} unmountOnClose={!retain}
      data-stage-demo="surface" data-stage-instance={instance}>
      <Editor.Root methods={methods} onSubmit={save} data-stage-demo="form"
        className={sx(styles.form, flex.column, flexItem.grow, minHeight.zero, width.full, gap.medium)}>
        <Stage.Header data-stage-demo="header" title={{ ref: titleSlotRef, 'data-stage-demo': 'title-slot', children: jsx ? <em data-stage-demo="jsx-title">{titleText}</em> : titleText }}
          description={description ? (jsx ? <span data-stage-demo="jsx-description">Editable <strong>profile</strong></span> : descriptionText) : undefined}
          icon={icon ? { children: (Tag, props) => <Tag {...props}><Info24Regular data-stage-demo="callback-icon" width={80} height={80} /></Tag> } : undefined} showCloseButton={closeButton} closeLabel={`Close ${instance}`}
          action={{ children: (Tag, props) => <Tag {...props} data-stage-demo="action-slot">
            <Button type="button" className={sx(minWidth.zero)} icon={<ArrowRight20Regular />} data-stage-demo="header-action" onClick={() => setActionCount(value => value + 1)}>
              <span className={sx(minWidth.zero, text.truncate)}>Share</span>
            </Button>
          </Tag> }} />
        <Stage.Body data-stage-demo="body">
          <div className={sx(flex.column, gap.medium)}>
            <Editor.Input name="name" label="Document title" data-stage-demo="name"
              className={narrowField} root={{ className: controlWidth }} input={{ className: sx(minWidth.zero) }} />
            <div data-stage-demo="combobox"><Editor.Combobox name="category" label="Category"
              className={narrowField} root={{ className: controlWidth }} input={{ className: sx(minWidth.zero) }}
              options={[{ value: 'one', label: 'Report' }, { value: 'two', label: 'Project' }]} /></div>
            <output data-stage-demo="dirty">Dirty: {String(isDirty)}</output>
            <output data-stage-demo="pending">Pending: {String(isSubmitting)}</output>
            {/* The native modal focus boundary requires the completion simulation inside its surface. */}
            {isSubmitting && <Button type="button" data-stage-demo="complete-save" onClick={() => pendingSave.current?.()}>Complete save</Button>}
            <output data-stage-demo="action-count">Header actions: {actionCount}</output>
            {discard && <div role="status" data-stage-demo="discard-prompt">
              <p>Discard unsaved changes?</p>
              <Button type="button" data-stage-demo="discard" disabled={isSubmitting} onClick={() => {
                methods.reset(); setDiscard(false); setOpen(false); setStatus('Discarded changes');
              }}>Discard and close</Button>
              <Button type="button" data-stage-demo="keep-editing" onClick={() => setDiscard(false)}>Keep editing</Button>
            </div>}
            {!footer && <Button type="button" data-stage-demo="body-close" onClick={requestClose}>Close editor</Button>}
            {longBody && Array.from({ length: 35 }, (_, index) => <p key={index}>Scrollable profile note {index + 1}. Header and footer remain available while this body scrolls.</p>)}
          </div>
        </Stage.Body>
        {footer && <Stage.Footer data-stage-demo="footer" className={sx(flex.row, flex.wrap, alignItems.center, justifyContent.end, gap.small)}>
          <Editor.Submit data-stage-demo="save" appearance="primary">Save</Editor.Submit>
          <Button type="button" data-stage-demo="footer-close" onClick={requestClose}>Close editor</Button>
        </Stage.Footer>}
      </Editor.Root>
    </Stage>
    </DemoCard>
  </section>;
};

const StagePanel: React.FC = () => <div data-stage-demo="panel"><StageExample instance="A" /></div>;
export default StagePanel;
