import * as React from 'react';
import {
  Button, Card, FluentProvider, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger,
  makeStyles, webDarkTheme, webLightTheme,
} from '@fluentui/react-components';
import { datePickerClassNames } from '@fluentui/react-datepicker-compat';
import { Info24Regular } from '@fluentui/react-icons';
import {
  Stage, alignItems, createForm, flex, flexItem, gap, justifyContent, minHeight, minWidth, useSx, width,
} from '@apvee/spfx-react-toolkit';
import type {
  StageDismissMode, StageOpenChangeData, StageOpenChangeEvent, StagePresentation, StageProps, StageSize,
} from '@apvee/spfx-react-toolkit';

interface Draft {
  name: string;
  category: string;
  // Form.DatePicker preserves the compat control's nullable empty-selection domain.
  // eslint-disable-next-line @rushstack/no-new-null
  date: Date | null;
}
const Editor = createForm<Draft>();
// Compat DatePicker requires null as its empty date value.
// eslint-disable-next-line @rushstack/no-new-null
const defaults: Draft = { name: 'Ada', category: 'one', date: null };
const useStyles = makeStyles({
  instance: { minInlineSize: '0', padding: '12px' },
  form: { alignSelf: 'stretch', overflow: 'hidden' },
  output: { overflowWrap: 'anywhere' },
  narrowField: {
    // Form.DatePicker exposes className on Field rather than a native root slot.
    // Scope the public compat root class to this Field so its input and calendar
    // affordance fit the same narrow body as the other controls.
    [`& > .${datePickerClassNames.root}`]: { minWidth: '0', width: '100%', boxSizing: 'border-box' },
  },
});

/** Local state deliberately belongs to the surface to make retention observable. */
const SurfaceCounter: React.FC<{ onMount: () => void }> = ({ onMount }) => {
  const [count, setCount] = React.useState(0);
  React.useEffect(onMount, [onMount]);
  return <div>
    <Button data-stage-demo="increment" onClick={() => setCount(value => value + 1)}>Increment surface counter</Button>
    <output data-stage-demo="counter">{count}</output>
  </div>;
};

const StageExample: React.FC<{ instance: string }> = ({ instance }) => {
  const sx = useSx();
  const styles = useStyles();
  const controlWidth = sx(minWidth.zero, width.full);
  const narrowField = sx(styles.narrowField, minWidth.zero, width.full);
  const methods = Editor.useForm({ defaultValues: defaults });
  // Read during render to subscribe to RHF's formState proxy updates.
  const { isDirty, isSubmitting } = methods.formState;
  const [presentation, setPresentation] = React.useState<StagePresentation>('dialog');
  const [size, setSize] = React.useState<StageSize>('small');
  const [open, setOpen] = React.useState(false);
  const [longBody, setLongBody] = React.useState(false);
  const [jsx, setJsx] = React.useState(false);
  const [icon, setIcon] = React.useState(true);
  const [description, setDescription] = React.useState(true);
  const [closeButton, setCloseButton] = React.useState(true);
  const [footer, setFooter] = React.useState(true);
  const [retain, setRetain] = React.useState(false);
  const [dismissMode, setDismissMode] = React.useState<StageDismissMode>('standard');
  const [inlineMenu, setInlineMenu] = React.useState(false);
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
  const onMount = React.useCallback(() => setMountCount(value => value + 1), []);
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
  const trigger = <Button data-stage-demo="trigger" appearance="primary">Open {instance}</Button>;
  const surfaceOptions: Extract<StageProps, { open: boolean }> = presentation === 'popover'
    ? { presentation, size: size === 'full' ? 'small' : size, withArrow, trigger, open, children: undefined }
    : { presentation, size, trigger, open, children: undefined };

  return <section data-stage-demo="instance" data-stage-instance={instance} className={sx(styles.instance, flex.column, gap.medium)}>
    <h3>Stage {instance}</h3>
    <div className={sx(flex.row, flex.wrap, gap.medium)}>
      <label>Presentation <select data-stage-demo="mode" value={presentation} disabled={open} onChange={event => changePresentation(event.target.value)}>
        <option value="dialog">Dialog</option><option value="drawer">Drawer</option><option value="popover">Popover</option>
      </select></label>
      <label>Size <select data-stage-demo="size" value={size} onChange={event => {
        const value = event.target.value;
        if (value === 'small' || value === 'medium' || value === 'large' || (value === 'full' && presentation !== 'popover')) setSize(value);
      }}>
        <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
        {presentation !== 'popover' && <option value="full">Full</option>}
      </select></label>
      <label>Dismiss <select data-stage-demo="dismiss" value={dismissMode} onChange={event => setDismissMode(event.target.value === 'explicit' ? 'explicit' : 'standard')}>
        <option value="standard">Standard</option><option value="explicit">Explicit</option>
      </select></label>
      <label><input data-stage-demo="long-body" type="checkbox" checked={longBody} onChange={event => setLongBody(event.target.checked)} /> Long body</label>
      <label><input data-stage-demo="jsx" type="checkbox" checked={jsx} onChange={event => setJsx(event.target.checked)} /> JSX title/description</label>
      <label><input data-stage-demo="icon" type="checkbox" checked={icon} onChange={event => setIcon(event.target.checked)} /> Icon</label>
      <label><input data-stage-demo="description" type="checkbox" checked={description} onChange={event => setDescription(event.target.checked)} /> Description</label>
      <label><input data-stage-demo="close" type="checkbox" checked={closeButton} onChange={event => setCloseButton(event.target.checked)} /> Header X</label>
      <label><input data-stage-demo="footer-toggle" type="checkbox" checked={footer} onChange={event => setFooter(event.target.checked)} /> Footer</label>
      <label><input data-stage-demo="retention" type="checkbox" checked={retain} onChange={event => setRetain(event.target.checked)} /> Retain after close</label>
      <label><input data-stage-demo="menu-inline" type="checkbox" checked={inlineMenu} onChange={event => setInlineMenu(event.target.checked)} /> Inline menu</label>
      <label><input data-stage-demo="dirty-guard" type="checkbox" checked={guardDirty} onChange={event => setGuardDirty(event.target.checked)} /> Guard dirty close</label>
      <label><input data-stage-demo="hold-save" type="checkbox" checked={holdSave} onChange={event => setHoldSave(event.target.checked)} /> Hold save until completed</label>
    </div>
    <p>Presentation is fixed for the open session. Configure the next session while closed. The same form content is used in every mode.</p>
    <Stage {...surfaceOptions} open={open} onOpenChange={onOpenChange} dismissMode={dismissMode} unmountOnClose={!retain}
      data-stage-demo="surface" data-stage-instance={instance}>
      <Editor.Root methods={methods} onSubmit={save} data-stage-demo="form"
        className={sx(styles.form, flex.column, flexItem.grow, minHeight.zero, width.full, gap.medium)}>
        <Stage.Header data-stage-demo="header" title={jsx ? <em data-stage-demo="jsx-title">Profile {instance}</em> : `Profile ${instance}`}
          description={description ? (jsx ? <span data-stage-demo="jsx-description">Editable <strong>profile</strong></span> : 'Edit profile details') : undefined}
          icon={icon ? <Info24Regular /> : undefined} showCloseButton={closeButton} closeLabel={`Close ${instance}`}
          actions={<>
            <Button data-stage-demo="header-action" onClick={() => setActionCount(value => value + 1)}>Header action</Button>
            {presentation === 'popover' && <Button data-stage-demo="arrow-toggle" onClick={() => setWithArrow(value => !value)}>
              {withArrow ? 'Hide arrow' : 'Show arrow'}
            </Button>}
          </>} />
        <Stage.Body data-stage-demo="body">
          <div className={sx(flex.column, gap.medium)}>
            <Editor.Input name="name" label={`Name ${instance}`} data-stage-demo="name"
              className={narrowField} root={{ className: controlWidth }} input={{ className: sx(minWidth.zero) }} />
            <div data-stage-demo="combobox"><Editor.Combobox name="category" label={`Category ${instance}`}
              className={narrowField} root={{ className: controlWidth }} input={{ className: sx(minWidth.zero) }}
              options={[{ value: 'one', label: 'One' }, { value: 'two', label: 'Two' }]} /></div>
            <div data-stage-demo="date-picker"><Editor.DatePicker name="date" label={`Date ${instance}`} className={narrowField} /></div>
            <Menu inline={inlineMenu}>
              <MenuTrigger disableButtonEnhancement><Button data-stage-demo="nested-menu">Nested menu</Button></MenuTrigger>
              <MenuPopover><MenuList><MenuItem onClick={() => setStatus('Nested menu action')}>Menu action</MenuItem></MenuList></MenuPopover>
            </Menu>
            <SurfaceCounter onMount={onMount} />
            <output data-stage-demo="dirty">Dirty: {String(isDirty)}</output>
            <output data-stage-demo="pending">Pending: {String(isSubmitting)}</output>
            <output data-stage-demo="action-count">Header actions: {actionCount}</output>
            {isSubmitting && <Button data-stage-demo="complete-save" onClick={() => pendingSave.current?.()}>Complete save</Button>}
            {discard && <div role="status" data-stage-demo="discard-prompt">
              <p>Discard unsaved changes?</p>
              <Button data-stage-demo="discard" disabled={isSubmitting} onClick={() => {
                methods.reset(); setDiscard(false); setOpen(false); setStatus('Discarded changes');
              }}>Discard and close</Button>
              <Button data-stage-demo="keep-editing" onClick={() => setDiscard(false)}>Keep editing</Button>
            </div>}
            {!footer && <Button data-stage-demo="body-close" onClick={requestClose}>Close editor</Button>}
            {longBody && Array.from({ length: 35 }, (_, index) => <p key={index}>Scrollable profile note {index + 1}. Header and footer remain available while this body scrolls.</p>)}
          </div>
        </Stage.Body>
        {footer && <Stage.Footer data-stage-demo="footer" className={sx(flex.row, flex.wrap, alignItems.center, justifyContent.end, gap.small)}>
          <Editor.Submit data-stage-demo="save" appearance="primary">Save</Editor.Submit>
          <Editor.Reset data-stage-demo="reset" onClick={() => { setDiscard(false); setStatus('Reset changes'); }}>Reset</Editor.Reset>
          <Button data-stage-demo="footer-close" onClick={requestClose}>Close editor</Button>
        </Stage.Footer>}
      </Editor.Root>
    </Stage>
    <output data-stage-demo="request" className={styles.output}>{request}</output>
    <output data-stage-demo="status">{status}</output>
    <output data-stage-demo="save-count">Save callbacks: {saveCount}</output>
    <output data-stage-demo="mount-count">Surface mounts: {mountCount}</output>
    <output data-stage-demo="open">Open: {String(open)}</output>
  </section>;
};

const StagePanel: React.FC = () => {
  const sx = useSx();
  const [selected, setSelected] = React.useState(false);
  return <div data-stage-demo="panel" className={sx(flex.column, gap.large)}>
    <h2>Stage — explicit surfaces, forms and independent providers</h2>
    <FluentProvider theme={webLightTheme} dir="ltr"><StageExample instance="A" /></FluentProvider>
    <FluentProvider theme={webDarkTheme} dir="rtl">
      <Card data-stage-demo="card" selected={selected} onSelectionChange={(_, data) => setSelected(data.selected)}>
        <p>Selectable ancestor — Stage heading must label its own surface independently.</p>
        <output data-stage-demo="card-selection">Card selected: {String(selected)}</output>
        <StageExample instance="B" />
      </Card>
    </FluentProvider>
  </div>;
};
export default StagePanel;
