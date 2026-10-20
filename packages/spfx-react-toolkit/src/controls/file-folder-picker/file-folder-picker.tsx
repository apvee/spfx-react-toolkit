import * as React from 'react';
import { Avatar, Button, Field, Tag, TagGroup, Tooltip, mergeClasses, useInputStyles_unstable } from '@fluentui/react-components';
import { resolveShorthand } from '@fluentui/react-utilities';
import type { FieldControlProps, InputState } from '@fluentui/react-components';
import { AttachRegular, DocumentRegular, DocumentPdfRegular, FolderDocumentRegular, FolderOpenRegular, FolderRegular } from '@fluentui/react-icons';
import { Stage } from '../stage/stage';
import { PickerControlBoundaryContext } from '../pickers/picker-runtime-context.internal';
import { fileFolderPickerIdentity } from './file-folder-picker-identity.internal';
import { useFileFolderPicker } from './useFileFolderPicker.internal';
import { useFileFolderPickerStyles } from './file-folder-picker.styles.internal';
import type { FileFolderPickerProps } from './file-folder-picker.types';

/** Select existing SharePoint/OneDrive Business files or folders through Microsoft's hosted v8 UI.
 * Requires a matching SPFx provider and existing delegated permissions. Only explicit accepted actions emit onChange.
 * Uses native Fluent Field slots; the remote Microsoft surface does not inherit arbitrary host CSS.
 * No persistence or host-side file-byte download is performed. Native upload/create are explicit opt-ins with immediate effects.
 * @param props - Source restrictions, JSON selection ownership and native Field contracts. Uses a single responsive 1080×720 hosted dialog.
 * @returns Owned Field with one Fluent Input-styled tag surface, mode-specific opening action and optional Stage-hosted iframe.
 * @example
 * ```tsx
 * <FileFolderPicker label="Documents" defaultValue={[]} multiselect extensions={['.pdf']}
 *   scope={{ kind: 'siteCollection' }} onChange={data => console.log(data.values)} />
 * ```
 */
export function FileFolderPicker(props: FileFolderPickerProps): React.ReactElement {
  const styles = useFileFolderPickerStyles(), runtime = useFileFolderPicker(props);
  const Boundary = React.useContext(PickerControlBoundaryContext);
  const mode = props.itemTypes ?? 'files';
  const browseLabel = props.browseLabel ?? (mode === 'folders' ? 'Sfoglia cartelle…' : mode === 'all' ? 'Sfoglia file e cartelle…' : 'Sfoglia file…');
  const placeholder = mode === 'folders' ? 'Seleziona cartelle…' : mode === 'all' ? 'Seleziona file o cartelle…' : 'Seleziona file…';
  const BrowseIcon = mode === 'folders' ? FolderOpenRegular : mode === 'all' ? FolderDocumentRegular : AttachRegular;
  const disabled = props.disabled || !runtime.ready;
  const validationState = props.validationState ?? (runtime.validation || runtime.error || runtime.initError ? 'error' : undefined);
  // Borrow only the native Input style state: the surface remains a group of real buttons/tags,
  // with no hidden/read-only input or editable-text semantics. Fluent owns exact border states,
  // dimensions, contrast and focus animation; only tag layout and the open-session underline extend it.
  const inputStyleState: InputState = {
    components: { root: 'span', input: 'input', contentBefore: 'span', contentAfter: 'span' },
    size: props.size ?? 'medium', appearance: 'outline',
    root: {},
    input: { disabled: props.disabled, 'aria-invalid': validationState === 'error' },
  };
  useInputStyles_unstable(inputStyleState);
  const content = (fieldProps: FieldControlProps): React.ReactElement => <div className={styles.region}>
    <div role="group" aria-label={props['aria-label'] ?? 'Selezione file e cartelle'} aria-busy={runtime.busy} aria-disabled={props.disabled || undefined}
      className={mergeClasses(inputStyleState.root.className, styles.control, runtime.busy && !props.disabled && styles.open, props.controlClassName)}>
    <div className={styles.selection}>
    {runtime.values.length === 0 ? undefined : <TagGroup className={styles.tags} size={props.size === 'large' ? 'small' : 'extra-small'} disabled={props.disabled} aria-label={props['aria-label'] ?? 'Selezione file e cartelle'} onDismiss={(_, data) => {
      const value = runtime.values.find(candidate => fileFolderPickerIdentity(candidate) === data.value);
      if (value !== undefined) runtime.remove(value);
    }}>
      {runtime.values.map(value => {
        const text = runtime.label(value);
        const Icon = value.kind === 'folder' ? FolderRegular : /^\.?pdf$/i.test(runtime.extension?.(value) ?? value.name?.split('.').pop() ?? '') ? DocumentPdfRegular : DocumentRegular;
        return <Tag key={fileFolderPickerIdentity(value)} value={fileFolderPickerIdentity(value)} dismissible disabled={props.disabled} className={styles.tag}
          primaryText={{ className: styles.text, title: text }} media={<Avatar icon={<Icon />} aria-label={value.kind === 'folder' ? 'Cartella' : 'File'} />} dismissIcon={{ 'aria-label': `Rimuovi ${text}` }}>{text}</Tag>;
      })}
    </TagGroup>}
    {runtime.values.length === 0 ? <button {...fieldProps} id={undefined} type="button" className={styles.placeholder} disabled={disabled || runtime.busy} onClick={runtime.browse}
      aria-label={props['aria-label']} aria-labelledby={props['aria-labelledby'] ?? fieldProps['aria-labelledby']}>{placeholder}</button> : undefined}
    </div>
    <Tooltip content={browseLabel} relationship="description">
      <Button {...fieldProps} type="button" id={props.id ?? fieldProps.id} ref={runtime.browseButton} className={styles.action} appearance="transparent" size={props.size === 'large' ? 'medium' : 'small'} disabled={disabled || runtime.busy} onClick={runtime.browse}
        aria-haspopup="dialog" aria-label={props['aria-label'] ?? browseLabel} aria-labelledby={props['aria-labelledby'] ?? fieldProps['aria-labelledby']}
        icon={<BrowseIcon />} />
    </Tooltip>
    </div>
  </div>;
  return <>
    <Field className={mergeClasses(styles.root, props.className)} label={props.id === undefined ? props.label : resolveShorthand(props.label, { defaultProps: { htmlFor: props.id } })} hint={props.hint} required={props.required}
      validationState={validationState} validationMessage={props.validationMessage ?? runtime.validation ?? runtime.error ?? runtime.initError}>
      {fieldProps => Boundary === undefined ? content(fieldProps) : <Boundary pending={runtime.hydrating}>{content(fieldProps)}</Boundary>}
    </Field>
    <Stage presentation="dialog" size="large" className={styles.dialog} open={runtime.dialog !== undefined} onOpenChange={(_, data) => { if (!data.open) runtime.cancel(); }}>
      <Stage.Header title={browseLabel} showCloseButton={false} className={styles.hiddenHeader} />
      <Stage.Body className={styles.dialogBody}>{runtime.dialog === undefined ? undefined : <iframe key={runtime.dialog.name} name={runtime.dialog.name} ref={runtime.attachFrame} title={browseLabel} className={styles.frame} />}</Stage.Body>
    </Stage>
  </>;
}
