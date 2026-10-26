import * as React from 'react';
import { EditableText } from '../../editable-text/editable-text';
import type { EditableTextBlock, EditableTextChangeData } from '../../editable-text/editable-text.types';
import { resolveWebPartShellStrings } from '../localization/web-part-shell-locales.internal';

/** Controlled private title boundary; only the Shell owns persistence. */
export interface WebPartTitleProps {
  title: string;
  titleHTML?: string;
  showTitle: boolean;
  editing: boolean;
  locale?: string;
  onChange: (update: { title: string; titleHTML: string }) => void;
}

const titleBlocks: readonly EditableTextBlock[] = ['h1', 'h2', 'h3', 'h4', 'p'];

/** Adapt title persistence and localized naming to the common text control. */
export function WebPartTitle(props: WebPartTitleProps): React.ReactElement {
  const strings = resolveWebPartShellStrings(props.locale);
  const publish = React.useCallback((data: EditableTextChangeData): void => {
    props.onChange({ title: data.text, titleHTML: data.html });
  }, [props.onChange]);
  if (!props.showTitle) return <></>;
  return <EditableText value={props.titleHTML} fallbackText={props.title}
    mode={props.editing ? 'edit' : 'read'} onChange={publish}
    allowedBlocks={titleBlocks} defaultBlock="h2" locale={props.locale}
    aria-label={strings.titleEditorLabel} placeholder={strings.titlePlaceholder} />;
}
