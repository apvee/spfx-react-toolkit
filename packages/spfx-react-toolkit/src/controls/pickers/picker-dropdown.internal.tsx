import * as React from 'react';
import { Dropdown, SearchBox, mergeClasses, useId } from '@fluentui/react-components';
import type { DropdownProps, Listbox } from '@fluentui/react-components';
import { useMergedRefs } from '@fluentui/react-utilities';
import type { PickerSelectionEvent } from './picker.types';
import { usePickerStyles } from './picker-styles.internal';

type KeyboardHandler = React.KeyboardEventHandler<HTMLElement>;
type NativeListboxProps = React.ComponentPropsWithRef<typeof Listbox>;
interface PickerDropdownProps {
  readonly fieldProps: Pick<DropdownProps, 'id' | 'aria-label' | 'aria-labelledby' | 'aria-describedby' | 'aria-invalid' | 'aria-required'>;
  readonly runtime: {
    readonly open: boolean; readonly query: string; readonly confirming?: boolean; readonly selectedOptions: readonly string[];
    setOpen(open: boolean): void;
    setQuery(query: string): void;
    select(event: PickerSelectionEvent, key: string): Promise<boolean>;
    remove(event: PickerSelectionEvent, key: string): Promise<void>;
    clear(event: PickerSelectionEvent): Promise<void>;
  };
  readonly multiselect: boolean;
  readonly disabled?: boolean;
  readonly searchable: boolean;
  readonly size?: 'medium' | 'large';
  readonly className: string;
  readonly selectionText: string;
  readonly selectionMedia?: React.ReactNode;
  readonly searchPlaceholder: string;
  readonly name?: string;
  // Native DOM FocusEvent.relatedTarget uses null when the next focus target is unavailable.
  readonly contains: (target: EventTarget | null | undefined) => boolean;
  readonly popupRef: React.Ref<HTMLDivElement>;
  readonly children: React.ReactNode;
}

/** Binds the completed public native handler without replacing Fluent navigation. */
function DropdownTrigger({ nativeProps, bind, media, className, textClassName }: {
  nativeProps: React.ComponentPropsWithRef<'button'>; bind: (handler: KeyboardHandler | undefined) => void;
  media: React.ReactNode; className: string; textClassName: string;
}): React.ReactElement {
  React.useLayoutEffect(() => { bind(nativeProps.onKeyDown); return () => bind(undefined); }, [bind, nativeProps.onKeyDown]);
  const decorateLabel = (children: React.ReactNode): React.ReactNode => React.Children.map(children, child => {
    // Fluent's public render slot may group completed children in a Fragment.
    if (React.isValidElement<{ children?: React.ReactNode }>(child) && child.type === React.Fragment) {
      return React.cloneElement(child, undefined, decorateLabel(child.props.children));
    }
    // Our controlled value/placeholder is plain text; native element siblings stay untouched.
    return typeof child === 'string' || typeof child === 'number' ?
      <span className={className}>{media}<span className={textClassName}>{child}</span></span> : child;
  });
  return <button {...nativeProps}>{decorateLabel(nativeProps.children)}</button>;
}

/** Search is a sibling of the semantic native list; native options retain their handlers and IDs. */
function DropdownSurface({ nativeProps, NativeListbox, popupRef, className, listClassName, children }: {
  nativeProps: NativeListboxProps; NativeListbox: React.ElementType<NativeListboxProps>; popupRef: React.Ref<HTMLDivElement>;
  className: string; listClassName: string; children: React.ReactNode;
}): React.ReactElement {
  const styles = usePickerStyles();
  const { ref, className: nativeClassName, style, ...listProps } = nativeProps;
  const mergedRef = useMergedRefs(ref, popupRef);
  // Fluent's positioner owns this style/ref. Forward them unchanged: classes cannot replace its runtime coordinates.
  return <div ref={mergedRef} className={mergeClasses(nativeClassName, className)} style={style}>
    <div className={styles.surfaceContent}>{children}<NativeListbox {...listProps} className={listClassName} /></div>
  </div>;
}

/** Private Dropdown presentation for the historical `combobox` API value. */
export function PickerDropdown({ fieldProps, runtime, multiselect, disabled, searchable, size, className, selectionText, selectionMedia,
  searchPlaceholder, name, contains, popupRef, children }: PickerDropdownProps): React.ReactElement {
  const styles = usePickerStyles(), searchId = useId('picker-search');
  // React/Fluent DOM ref contracts use null for detached nodes.
  const triggerRef = React.useRef<HTMLButtonElement>(null), searchRef = React.useRef<HTMLInputElement>(null);
  const nativeKeyDown = React.useRef<KeyboardHandler>();
  const [activeId, setActiveId] = React.useState<string>();
  const session = React.useRef({ open: runtime.open, version: 0 });
  React.useLayoutEffect(() => {
    if (session.current.open !== runtime.open) session.current = { open: runtime.open, version: session.current.version + 1 };
  }, [runtime.open]);
  const setOpen = (open: boolean): void => {
    if (session.current.open !== open) session.current = { open, version: session.current.version + 1 };
    runtime.setOpen(open);
  };
  const bind = React.useCallback((handler: KeyboardHandler | undefined) => { nativeKeyDown.current = handler; }, []);
  React.useEffect(() => () => { session.current = { open: false, version: session.current.version + 1 }; }, []);
  React.useEffect(() => {
    if (!runtime.open) { setActiveId(undefined); return; }
    const version = session.current.version, owner = triggerRef.current?.ownerDocument;
    const frame = owner?.defaultView?.requestAnimationFrame(() => {
      if (session.current.open && session.current.version === version && (contains(owner.activeElement) || owner.activeElement === owner.body)) searchRef.current?.focus();
    });
    return () => { if (frame !== undefined) owner?.defaultView?.cancelAnimationFrame(frame); };
  }, [runtime.open]);
  const closeFromKeyboard = (event: React.KeyboardEvent<HTMLElement>, preventDefault: boolean): void => {
    event.stopPropagation();
    if (preventDefault) event.preventDefault();
    // Tab continues its normal browser step from the trigger, avoiding native implicit selection.
    triggerRef.current?.focus(); setOpen(false);
  };
  const select: NonNullable<DropdownProps['onOptionSelect']> = async (event, data) => {
    if (disabled || runtime.confirming) return;
    const key = data.optionValue;
    if (key === undefined) { if (data.selectedOptions.length === 0) await runtime.clear(event); return; }
    if (runtime.selectedOptions.includes(key)) { if (multiselect) await runtime.remove(event, key); return; }
    const version = session.current.version;
    const accepted = await runtime.select(event, key);
    if (accepted && !multiselect && session.current.open && version === session.current.version) {
      const restore = contains(triggerRef.current?.ownerDocument.activeElement);
      setOpen(false); if (restore) triggerRef.current?.focus();
    }
  };
  return <Dropdown {...fieldProps} disabled={disabled} aria-disabled={disabled || runtime.confirming || undefined} size={size} className={className}
    // Native viewport sizing constrains the outer surface; the inner content retains the toolkit height cap.
    positioning={{ autoSize: true }} value={selectionText} placeholder={multiselect ? 'Seleziona elementi' : 'Seleziona un elemento'}
    multiselect={multiselect} selectedOptions={[...runtime.selectedOptions]} clearable={!multiselect} open={runtime.open}
    onOpenChange={(event, data) => {
      if (!data.open) {
        // The shared logical boundary owns blur dismissal after native portal focus transfers settle.
        if (event.type === 'blur') return;
        const target = event.target, owner = triggerRef.current?.ownerDocument.defaultView;
        if (event.type === 'click' && owner && target instanceof owner.Element && contains(target) && target.closest('[role="option"], [role="menuitemcheckbox"]')) return;
        if (event.type === 'keydown' && 'key' in event && (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowUp' && event.altKey)) return;
      }
      if (data.open && runtime.confirming) return;
      setOpen(data.open);
    }} onOptionSelect={select} onActiveOptionChange={(_, data) => setActiveId(data.nextOption?.id)}
    button={{ ref: triggerRef, onClickCapture: event => { if (runtime.confirming) { event.preventDefault(); event.stopPropagation(); } },
      onKeyDownCapture: event => {
        if (runtime.open && event.key === 'Tab') { closeFromKeyboard(event, false); return; }
        if (runtime.open && event.key === 'F2' && searchable) { event.preventDefault(); event.stopPropagation(); searchRef.current?.focus(); return; }
        if (runtime.confirming && event.key !== 'Escape' && event.key !== 'Tab') { event.preventDefault(); event.stopPropagation(); }
      }, children: (_, nativeProps) => <DropdownTrigger nativeProps={nativeProps} bind={bind} media={selectionMedia} className={styles.selection} textClassName={styles.text} /> }}
    clearButton={{ 'aria-label': 'Cancella selezione', disabled: runtime.confirming }}
    listbox={{ 'aria-label': name === undefined ? 'Opzioni' : `${name}: opzioni`, children: (NativeListbox, nativeProps) =>
      <DropdownSurface nativeProps={nativeProps} NativeListbox={NativeListbox} popupRef={popupRef}
        className={mergeClasses(styles.surface, !runtime.open && styles.hidden)} listClassName={styles.dropdownList}>
        {searchable ? <SearchBox id={searchId} required={false} aria-required={false} aria-invalid={false} aria-label={name === undefined ? 'Cerca opzioni' : `${name}: cerca`} className={styles.search} size={size} value={runtime.query} placeholder={searchPlaceholder}
          dismiss={{ 'aria-label': 'Cancella ricerca', 'aria-disabled': runtime.confirming || undefined, onClickCapture: event => { if (runtime.confirming) { event.preventDefault(); event.stopPropagation(); } } }}
          // Search owns its name; Field injects aria-labelledby even when an independent aria-label is provided.
          input={{ ref: searchRef, 'aria-labelledby': undefined, autoComplete: 'off', readOnly: runtime.confirming, 'aria-disabled': runtime.confirming || undefined,
            'aria-controls': nativeProps.id,
            'aria-describedby': fieldProps['aria-describedby'], 'aria-activedescendant': runtime.open ? activeId : undefined }}
          onChange={(_, data) => { setActiveId(undefined); runtime.setQuery(data.value); }}
          onKeyDown={event => {
            if (event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229) return;
            if (event.key === 'Escape') { closeFromKeyboard(event, true); return; }
            if (event.key === 'Tab') { closeFromKeyboard(event, false); return; }
            if (runtime.confirming || event.ctrlKey || event.metaKey || event.altKey || runtime.query && ['Home', 'End'].includes(event.key)) return;
            if (['ArrowDown', 'ArrowUp', 'Home', 'End', 'PageDown', 'PageUp', 'Enter'].includes(event.key)) {
              nativeKeyDown.current?.(event); event.stopPropagation();
            }
          }} /> : undefined}
      </DropdownSurface> }}>
    {children}
  </Dropdown>;
}
