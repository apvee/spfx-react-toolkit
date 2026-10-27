import * as React from 'react';
import type { ChainedCommands, Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { SelectionBookmark } from '@tiptap/pm/state';
import { makeStyles } from '@griffel/react';
import { elementContains } from '@fluentui/react-utilities';
import { tokens } from '@fluentui/react-theme';
import { ThemeContext_unstable } from '@fluentui/react-shared-contexts';
import { Toolbar, Button, ToolbarToggleButton, ToolbarDivider, Dropdown, Combobox, Option,
  Popover, PopoverSurface, PopoverTrigger, Tooltip,
  Menu, MenuTrigger, MenuPopover, MenuList, MenuItem, MenuItemRadio, MenuItemCheckbox, MenuDivider,
  Overflow, OverflowItem, useOverflowMenu, useIsOverflowItemVisible } from '@fluentui/react-components';
import { TextBold20Regular, TextItalic20Regular, TextUnderline20Regular,
  TextColor20Regular, TextClearFormatting20Regular, TextLineSpacing20Regular, MoreHorizontal20Regular } from '@fluentui/react-icons';
import { useStableCallback } from '../../hooks/useStableCallback';
import { useSx } from '../../hooks/useSx';
import * as width from '../../helpers/styles/width';
import { useResizeObserver } from '../../utils/resize-observer.internal';
import type { ContainerSize } from '../../core/types';
import { editableTextFormatRegistry, isEditableTextHeadingLevel } from './editable-text-format.internal';
import { clearEditableTextFormat, readEditableTextToolbarFormat, setEditableTextFontSize, setEditableTextColor, toggleEditableTextMark } from './editable-text-formatting.internal';
import type { EditableTextBlockPolicy } from './editable-text-policy.internal';
import { resolveEditableTextBlockPolicy } from './editable-text-policy.internal';
import type { EditableTextSelectionFormat } from './editable-text-formatting.internal';
import { EditableTextColorPicker } from './editable-text-color-picker.internal';
import { resolveEditableTextStrings } from './localization/editable-text-strings.internal';
import type { EditableTextStrings } from './editable-text.types';

const useToolbarStyles = makeStyles({
  surface: { padding: '2px', maxWidth: 'calc(100vw - 16px)', boxSizing: 'border-box' },
  toolbar: { minInlineSize: 0, inlineSize: '100%', padding: 0, flexWrap: 'nowrap', alignItems: 'center',
    '& > *': { flexShrink: 0 },
    '& [data-overflowing]': { display: 'none' } },
  // Fluent slots set physical min-width:250px; override that same property
  // so their compact controls fit the measured overflow toolbar in either dir.
  headingWrapper: { paddingInlineEnd: tokens.spacingHorizontalS },
  heading: { minWidth: '112px', width: '112px' },
  headingButton: { gridTemplateColumns: '[content] minmax(0, 1fr) [icon] auto [end]' },
  headingValue: { minInlineSize: 0, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  size: { minWidth: '64px', width: '64px', '& input': { minWidth: '0px', width: '100%' } },
  palette: { display: 'flex', gap: tokens.spacingHorizontalS },
});
interface EditableTextToolbarProps { editor: Editor; ownerDocument: Document; policy?: EditableTextBlockPolicy; strings?: EditableTextStrings }
interface ToolbarActions {
  run: (action: () => void) => void;
  state: EditableTextSelectionFormat;
  editor: Editor;
  strings: EditableTextStrings;
  policy: EditableTextBlockPolicy;
}

/** Prefix native block commands with a boundary on their shared transaction. */
function blockCommand(editor: Editor): ChainedCommands {
  return editor.chain().command(({ tr }) => { closeHistory(tr); return true; });
}
function SpacingItems({ run, editor, state, strings }: ToolbarActions): React.ReactElement {
  const change = (attrs: Record<string, number | true | undefined>): void => run(() => { blockCommand(editor).updateAttributes(state.block.type, attrs).run(); });
  return <MenuList>
    <MenuItemRadio persistOnClick={false} name="lineSpacing" value="default"
      onClick={() => change({ lineSpacing: undefined })}>{strings.defaultLineSpacing}</MenuItemRadio>
    {editableTextFormatRegistry.lineSpacing.map(line => <MenuItemRadio persistOnClick={false} key={line} name="lineSpacing" value={String(line)}
      onClick={() => change({ lineSpacing: line })}>{strings.lineSpacingValue(line)}</MenuItemRadio>)}
    <MenuDivider />
    <MenuItemCheckbox persistOnClick={false} name="spacing" value="before" onClick={() => change({ spaceBefore: state.block.spaceBefore ? undefined : true })}>{strings.spaceBefore}</MenuItemCheckbox>
    <MenuItemCheckbox persistOnClick={false} name="spacing" value="after" onClick={() => change({ spaceAfter: state.block.spaceAfter ? undefined : true })}>{strings.spaceAfter}</MenuItemCheckbox>
  </MenuList>;
}
function checkedSpacing(state: EditableTextSelectionFormat): Record<string, string[]> {
  return { lineSpacing: [state.block.lineSpacing === undefined ? 'default' : String(state.block.lineSpacing)],
    spacing: [...(state.block.spaceBefore ? ['before'] : []), ...(state.block.spaceAfter ? ['after'] : [])] };
}
/** Keeps a direct popup bound to its existing Fluent overflow trigger. */
function DirectColorCommand({ open, onOpenChange, run, state, editor, strings }: ToolbarActions & {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}): React.ReactElement {
  const colorVisible = useIsOverflowItemVisible('color');
  React.useLayoutEffect(() => {
    // Hiding the direct trigger must dismiss its portal and forget its open state;
    // otherwise narrowing leaves an orphan popup and widening resurrects it.
    if (!colorVisible) onOpenChange(false);
  }, [colorVisible, onOpenChange]);
  return <Popover open={open && colorVisible} onOpenChange={(_event, data) => onOpenChange(data.open)} positioning="below-start">
    <Tooltip content={state.color === 'mixed' ? strings.mixedFormatting(strings.themeColor) : strings.themeColor} relationship="description">
      <PopoverTrigger disableButtonEnhancement><Button appearance="subtle" size="medium" aria-label={strings.themeColor} icon={<TextColor20Regular />} /></PopoverTrigger>
    </Tooltip>
    <PopoverSurface><EditableTextColorPicker strings={strings} selected={state.color}
      onSelect={color => run(() => { setEditableTextColor(editor, color); })} /></PopoverSurface>
  </Popover>;
}

function OverflowCommands(props: ToolbarActions): React.ReactElement {
  const { ref, isOverflowing } = useOverflowMenu<HTMLButtonElement>();
  const headingVisible = useIsOverflowItemVisible('heading');
  const boldVisible = useIsOverflowItemVisible('bold');
  const italicVisible = useIsOverflowItemVisible('italic');
  const underlineVisible = useIsOverflowItemVisible('underline');
  const colorVisible = useIsOverflowItemVisible('color');
  const clearVisible = useIsOverflowItemVisible('clear');
  const spacingVisible = useIsOverflowItemVisible('spacing');
  const [colorOpen, setColorOpen] = React.useState(false);
  const [menuOpen, setMenuOpen] = React.useState(false);
  React.useLayoutEffect(() => {
    // Widening can remove More or just move color back to its direct trigger.
    // Dismiss the removed color subtree and its parent before a later narrowing
    // can mount that subtree again. Other valid More menus and inner Escape keep
    // their normal behavior while the corresponding trigger remains available.
    if (!isOverflowing || (colorVisible && colorOpen)) { setMenuOpen(false); setColorOpen(false); }
  }, [isOverflowing, colorVisible, colorOpen]);
  if (!isOverflowing) return <></>;
  const { run, editor, state, strings, policy } = props;
  return <Menu open={menuOpen} onOpenChange={(_event, data) => { setMenuOpen(data.open); if (!data.open) setColorOpen(false); }} checkedValues={{ marks: [state.bold === true ? 'bold' : '', state.italic === true ? 'italic' : '', state.underline === true ? 'underline' : ''] }}>
    <Tooltip content={strings.moreFormatting} relationship="description"><MenuTrigger disableButtonEnhancement><Button appearance="subtle" size="medium" ref={ref} aria-label={strings.moreFormatting} icon={<MoreHorizontal20Regular />} /></MenuTrigger></Tooltip>
    <MenuPopover><MenuList>
      {policy.blocks.length > 1 && !headingVisible && <Menu checkedValues={{ heading: [state.block.type === 'heading' ? String(state.block.level) : 'normal'] }}>
        <MenuTrigger disableButtonEnhancement><MenuItem>{strings.heading}</MenuItem></MenuTrigger>
        <MenuPopover><MenuList>{policy.blocks.filter(block => block !== 'p').map(block => Number(block.slice(1))).map(level => <MenuItemRadio key={level} name="heading" value={String(level)}
          onClick={() => run(() => { blockCommand(editor).setNode('heading', { level }).run(); })}>{strings.headingLevel(level)}</MenuItemRadio>)}
          {policy.blocks.includes('p') && <MenuItemRadio name="heading" value="normal" onClick={() => run(() => { blockCommand(editor).setNode('paragraph').run(); })}>{strings.normal}</MenuItemRadio>}
        </MenuList></MenuPopover>
      </Menu>}
      {(['bold', 'italic', 'underline'] as const).map(mark => {
        const visible = mark === 'bold' ? boldVisible : mark === 'italic' ? italicVisible : underlineVisible;
        if (visible) return <React.Fragment key={mark} />;
        const label = strings[mark];
        const toggle = (): void => run(() => {
          toggleEditableTextMark(editor, mark);
        });
        return state[mark] === 'mixed'
          ? <MenuItem key={mark} role="menuitemcheckbox" aria-checked="mixed" onClick={toggle}>{label}</MenuItem>
          : <MenuItemCheckbox key={mark} name="marks" value={mark} onClick={toggle}>{label}</MenuItemCheckbox>;
      })}
      {/* PopoverTrigger supplies button semantics even with enhancement disabled;
          retain the command's menu role explicitly at this integration boundary. */}
      {!colorVisible && <Popover open={colorOpen} onOpenChange={(_event, data) => setColorOpen(data.open)} positioning="below-start">
        <PopoverTrigger disableButtonEnhancement><MenuItem role="menuitem" persistOnClick>{strings.themeColor}</MenuItem></PopoverTrigger>
        <PopoverSurface><EditableTextColorPicker strings={strings} selected={state.color}
          onSelect={color => { run(() => { setEditableTextColor(editor, color); }); setColorOpen(false); setMenuOpen(false); }} /></PopoverSurface>
      </Popover>}
      {!clearVisible && <MenuItem onClick={() => run(() => { clearEditableTextFormat(editor); })}>{strings.clearFormatting}</MenuItem>}
      {!spacingVisible && <Menu checkedValues={checkedSpacing(state)}><MenuTrigger disableButtonEnhancement><MenuItem>{strings.spacing}</MenuItem></MenuTrigger>
        <MenuPopover><SpacingItems {...props} /></MenuPopover></Menu>}
    </MenuList></MenuPopover>
  </Menu>;
}

/** Private Fluent-only toolbar; the editor's existing publication callback is the sole writer. */
export function EditableTextToolbar({ editor, ownerDocument, strings = resolveEditableTextStrings(undefined), policy = resolveEditableTextBlockPolicy(undefined, undefined) }: EditableTextToolbarProps): React.ReactElement {
  const theme = React.useContext(ThemeContext_unstable);
  const readSelection = useStableCallback(() => readEditableTextToolbarFormat(editor, theme));
  const styles = useToolbarStyles();
  const sx = useSx();
  const [contentWidth, setContentWidth] = React.useState<number>();
  const [targetWidth, setTargetWidth] = React.useState<number>();
  const [viewportWidth, setViewportWidth] = React.useState<number>();
  const [visibleFootprint, setVisibleFootprint] = React.useState<{ budget: number; width: number }>();
  const [overflowRevision, setOverflowRevision] = React.useState(0);
  const [toolbarElement, setToolbarElement] = React.useState<HTMLDivElement>();
  const updateTargetWidth = React.useCallback((size: ContainerSize | undefined): void => {
    setTargetWidth(size?.width);
  }, []);
  useResizeObserver(editor.view.dom, updateTargetWidth);
  const updateViewportWidth = React.useCallback((size: ContainerSize | undefined): void => {
    setViewportWidth(size?.width);
  }, []);
  useResizeObserver(ownerDocument.documentElement, updateViewportWidth);
  const measureContent = React.useCallback((element: HTMLDivElement): void => {
    // Measure before Fluent hides any items, then retain that full footprint while
    // overflowed. Re-measuring hidden content would shrink the available width and
    // prevent native Overflow from recovering when the target becomes wider.
    if (element.querySelector('[data-overflowing]')) return;
    const view = element.ownerDocument.defaultView;
    const parent = element.parentElement;
    if (!view || !parent) return;
    const surfaceStyle = view.getComputedStyle(parent);
    const chrome = ['paddingLeft', 'paddingRight', 'borderLeftWidth', 'borderRightWidth'] as const;
    const natural = Array.from(element.children).filter(child => child.getAttribute('aria-hidden') !== 'true').reduce((total, child) => total + child.getBoundingClientRect().width, 0);
    if (natural > 0) setContentWidth(natural + chrome.reduce((total, property) => total + parseFloat(surfaceStyle[property] || '0'), 0));
  }, []);
  const attachToolbar = React.useCallback((element: HTMLDivElement | null): void => {
    // React's callback-ref boundary uses its required legacy null sentinel.
    setToolbarElement(element ?? undefined);
    if (element) { measureContent(element); setTargetWidth(editor.view.dom.getBoundingClientRect().width); }
  }, [editor, measureContent]);
  const availableWidth = contentWidth === undefined ? undefined : Math.min(contentWidth,
    targetWidth ?? contentWidth, viewportWidth === undefined ? contentWidth : Math.max(0, viewportWidth - 16));
  const surfaceWidth = visibleFootprint && visibleFootprint.budget === availableWidth ? visibleFootprint.width : availableWidth;
  const updateOverflow = React.useCallback((): void => { setOverflowRevision(revision => revision + 1); }, []);
  const fitVisibleContent = useStableCallback((): void => {
    if (!toolbarElement || availableWidth === undefined) return;
    // Wait for Fluent's More control to mount before measuring an overflowed row.
    // Its ref registration can change Fluent's final visibility decision.
    if (toolbarElement.querySelector('[data-overflowing]') && !toolbarElement.querySelector('[data-overflow-menu]')) return;
    const parent = toolbarElement.parentElement;
    const view = toolbarElement.ownerDocument.defaultView;
    if (!parent || !view) return;
    const css = view.getComputedStyle(parent);
    const chrome = ['paddingLeft', 'paddingRight', 'borderLeftWidth', 'borderRightWidth'] as const;
    const visible = Array.from(toolbarElement.children).filter(child => child.getAttribute('aria-hidden') !== 'true')
      .reduce((total, child) => total + child.getBoundingClientRect().width, 0);
    if (visible <= 0) return;
    const fitted = visible + chrome.reduce((total, property) => total + parseFloat(css[property] || '0'), 0);
    if (fitted > availableWidth + 1) return; // Fluent has not yet committed the smaller budget.
    setVisibleFootprint(current => current?.budget === availableWidth && current.width === fitted
      ? current : { budget: availableWidth, width: fitted });
  });
  const remeasureContent = React.useCallback((): void => {
    if (!toolbarElement) return;
    measureContent(toolbarElement);
    // Fluent observes this row during its layout effect, before our existing
    // observer subscribes. Its native resize callback queues overflow processing
    // as a microtask; fit afterward even if the resulting visibility mask is unchanged.
    // Growth therefore lets Fluent recover controls before measuring their width.
    queueMicrotask(fitVisibleContent);
  }, [toolbarElement, measureContent, fitVisibleContent]);
  useResizeObserver(toolbarElement, remeasureContent);
  React.useLayoutEffect(() => {
    // React has now committed Fluent's visibility and More control. The surface
    // fits that real row; the independent full-width budget remains unchanged.
    // A changed target/viewport budget first restores available width so native
    // Overflow can recover controls, then its next update fits the visible row.
    fitVisibleContent();
  }, [overflowRevision, toolbarElement, fitVisibleContent]);
  const [state, setState] = React.useState(() => readEditableTextToolbarFormat(editor, theme));
  React.useEffect(() => { setState(readSelection()); }, [theme, readSelection]);
  const [visible, setVisible] = React.useState(false);
  const [colorOpen, setColorOpen] = React.useState(false);
  const [sizeDraft, setSizeDraft] = React.useState<string>();
  const bookmark = React.useRef<SelectionBookmark>(editor.state.selection.getBookmark());
  const caretMarks = React.useRef(editor.state.storedMarks ?? undefined);
  const surface = React.useRef<HTMLDivElement>(null); // React17 DOM refs require its legacy null sentinel.
  React.useEffect(() => {
    const editorDom = editor.view.dom;
    const update = (): void => {
      // An authoritative replacement uses a fresh state and emits selectionUpdate.
      // Retain the current model range, never a stale range from another document.
      bookmark.current = editor.state.selection.getBookmark();
      caretMarks.current = editor.state.storedMarks ?? undefined;
      setState(readSelection()); setSizeDraft(undefined);
    };
    const focus = (event: FocusEvent): void => {
      const target = event.target;
      const view = ownerDocument.defaultView;
      if (!view || !(target instanceof view.Node)) return;
      if (editorDom.contains(target) || elementContains(surface.current, target)) setVisible(true);
      else setVisible(false);
    };
    const escape = (event: KeyboardEvent): void => {
      if (event.altKey && event.key === 'F10' && editorDom.contains(event.target as Node)) {
        event.preventDefault(); setVisible(true);
        queueMicrotask(() => {
          const controls = surface.current?.querySelectorAll<HTMLElement>('[role="toolbar"] button, [role="toolbar"] input');
          const first = controls && Array.from(controls).find(control => !control.closest('[data-overflowing]'));
          first?.focus();
        });
        return;
      }
      if (event.key !== 'Escape') return;
      if (editorDom.contains(event.target as Node)) { setVisible(false); editorDom.blur(); }
    };
    editor.on('transaction', update); editor.on('selectionUpdate', update);
    ownerDocument.addEventListener('focusin', focus);
    ownerDocument.addEventListener('keydown', escape, true);
    return () => {
      editor.off('transaction', update); editor.off('selectionUpdate', update);
      ownerDocument.removeEventListener('focusin', focus); ownerDocument.removeEventListener('keydown', escape, true);
    };
  }, [editor, ownerDocument, readSelection]);
  const restore = (): void => {
    if (editor.isDestroyed) return;
    const selection = bookmark.current.resolve(editor.state.doc);
    // Even setting an equal selection clears ProseMirror stored marks. Keep caret
    // overrides (including explicit []) intact while a Fluent control has focus.
    if (selection.eq(editor.state.selection)) return;
    const tr = editor.state.tr.setSelection(selection);
    if (selection.empty && caretMarks.current !== undefined) tr.setStoredMarks(caretMarks.current);
    editor.view.dispatch(tr);
  };
  const run = (action: () => void): void => {
    if (editor.isDestroyed || editor.view.composing) return;
    restore(); action(); setColorOpen(false); editor.view.focus();
  };
  const runPopupOption = (action: () => void): void => {
    if (editor.isDestroyed || editor.view.composing) return;
    run(action);
    const document = editor.state.doc;
    const selection = editor.state.selection;
    // Fluent listbox clicks focus the trigger while bubbling, after onOptionSelect.
    // Restore afterward only while this editor/document/range is still current.
    queueMicrotask(() => {
      if (!editor.isDestroyed && editor.state.doc === document && editor.state.selection.eq(selection)) editor.view.focus();
    });
  };
  const actions: ToolbarActions = { run, state, editor, strings, policy };
  const size = state.size === 'mixed' ? '' : String(state.size);
  const headingLabel = state.block.type === 'heading' ? strings.headingLevel(state.block.level ?? 2) : strings.normal;
  if (!visible || editor.isDestroyed) return <></>;
  // Default Fluent portal mounts inherit the Shell targetDocument and preserve
  // virtual-parent links for this toolbar's nested popups; body overrides do not.
  // V1 intentionally omits Format painter; commands edit only the current range.
  // Let Fluent position the content-sized surface and own overflow commands.
  return <Popover open={visible} unstable_disableAutoFocus trapFocus={false}
    positioning={{ target: editor.view.dom, position: 'above', align: 'start', offset: 4, autoSize: 'height', strategy: 'fixed' }}
    onOpenChange={(event, data) => {
      if (data.open) return;
      const view = ownerDocument.defaultView;
      const target = event.target;
      // Fluent's custom positioning target is not its triggerRef. Only pointer
      // dismissal from our editor/owned portals is inside; Escape/outside closes.
      const pointerDismissal = ['click', 'touchstart', 'contextmenu'].includes(event.type);
      if (pointerDismissal && view && target instanceof view.Node && !editor.isDestroyed &&
        (editor.view.dom.contains(target) || elementContains(surface.current, target))) return;
      setVisible(false);
    }}>
    <PopoverSurface ref={surface} className={sx(styles.surface, surfaceWidth !== undefined && width.px(surfaceWidth))} data-title-toolbar="" aria-label={strings.formattingControls}>
      <Overflow minimumVisible={1} padding={0} onOverflowChange={updateOverflow}>
        <Toolbar ref={attachToolbar} className={styles.toolbar} size="medium" aria-label={strings.formatting} checkedValues={{ marks: [state.bold === true ? 'bold' : '', state.italic === true ? 'italic' : '', state.underline === true ? 'underline' : ''] }}>
          {/* Fluent composite refs target trigger slots; measure/hide each complete control. */}
          {policy.blocks.length > 1 && <OverflowItem id="heading" priority={10}><div className={styles.headingWrapper}><Dropdown className={styles.heading} size="small" aria-label={strings.heading} value={headingLabel}
            button={{ className: styles.headingButton, children: <span className={styles.headingValue} title={headingLabel}>{headingLabel}</span> }}
            selectedOptions={[state.block.type === 'heading' ? String(state.block.level) : 'normal']} onOptionSelect={(_event, data) => runPopupOption(() => {
              const level = Number(data.optionValue);
              if (isEditableTextHeadingLevel(level) && policy.blocks.includes(`h${level}`)) blockCommand(editor).setNode('heading', { level }).run();
              else if (data.optionValue === 'normal' && policy.blocks.includes('p')) blockCommand(editor).setNode('paragraph').run();
            })}>
            {policy.blocks.filter(block => block !== 'p').map(block => Number(block.slice(1))).map(level => <Option key={level} value={String(level)}>{strings.headingLevel(level)}</Option>)}{policy.blocks.includes('p') && <Option value="normal">{strings.normal}</Option>}
          </Dropdown></div></OverflowItem>}
          <OverflowItem id="size" pinned><div><Combobox className={styles.size} size="small" aria-label={strings.fontSize} expandIcon={{ 'aria-label': strings.fontSize }} freeform value={sizeDraft ?? size}
            placeholder={state.size === 'mixed' ? strings.mixed : undefined} selectedOptions={state.size === 'mixed' ? [] : [size]}
            onChange={event => setSizeDraft(event.target.value)}
            onOptionSelect={(_event, data) => { if (data.optionValue) runPopupOption(() => { setEditableTextFontSize(editor, data.optionValue ?? ''); }); }}
            onKeyDown={event => {
              if (event.key === 'Enter' && sizeDraft !== undefined) { event.preventDefault(); const value = sizeDraft; run(() => { setEditableTextFontSize(editor, value); }); setSizeDraft(undefined); }
              else if (event.key === 'Escape') setSizeDraft(undefined);
            }} onBlur={() => setSizeDraft(undefined)}>
            {editableTextFormatRegistry.suggestedSizes.map(value => <Option key={value} value={String(value)}>{String(value)}</Option>)}
          </Combobox></div></OverflowItem>
          <OverflowItem id="marks-divider"><ToolbarDivider /></OverflowItem>
          {(['bold', 'italic', 'underline'] as const).map(mark => {
            const label = strings[mark];
            const icon = mark === 'bold' ? <TextBold20Regular /> : mark === 'italic' ? <TextItalic20Regular /> : <TextUnderline20Regular />;
            const toggle = (): void => run(() => {
              toggleEditableTextMark(editor, mark);
            });
            return <Tooltip key={mark} content={state[mark] === 'mixed' ? strings.mixedFormatting(label) : strings.formattingShortcut(label, `Ctrl/Cmd+${mark === 'bold' ? 'B' : mark === 'italic' ? 'I' : 'U'}`)} relationship="description">
              <OverflowItem id={mark}>{state[mark] === 'mixed'
                // Fluent ToggleButton owns its boolean aria-pressed. The ordinary
                // Fluent Button can accurately expose the third mixed state.
                ? <Button appearance="subtle" size="medium" aria-label={label} aria-pressed="mixed" icon={icon} onClick={toggle} />
                : <ToolbarToggleButton size="medium" name="marks" value={mark} aria-label={label} icon={icon} onClick={toggle} />}
              </OverflowItem>
            </Tooltip>;
          })}
          <OverflowItem id="color"><div><DirectColorCommand {...actions} open={colorOpen} onOpenChange={setColorOpen} /></div></OverflowItem>
          <OverflowItem id="clear"><Tooltip content={strings.clearFormatting} relationship="description"><Button appearance="subtle" size="medium" aria-label={strings.clearFormatting} icon={<TextClearFormatting20Regular />} onClick={() => run(() => { clearEditableTextFormat(editor); })} /></Tooltip></OverflowItem>
          <OverflowItem id="spacing-divider"><ToolbarDivider /></OverflowItem>
          <OverflowItem id="spacing"><div><Menu checkedValues={checkedSpacing(state)} positioning="below-start">
            <Tooltip content={strings.spacing} relationship="description"><MenuTrigger disableButtonEnhancement><Button appearance="subtle" size="medium" aria-label={strings.spacing} icon={<TextLineSpacing20Regular />} /></MenuTrigger></Tooltip>
            <MenuPopover><SpacingItems {...actions} /></MenuPopover>
          </Menu></div></OverflowItem>
          <OverflowCommands {...actions} />
        </Toolbar>
      </Overflow>
    </PopoverSurface>
  </Popover>;
}
