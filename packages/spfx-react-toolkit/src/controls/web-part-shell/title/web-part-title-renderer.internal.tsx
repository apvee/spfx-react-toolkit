import * as React from 'react';
import { isTitleDocument, isTitleDocumentEmpty } from './web-part-title-codec.internal';
import { getTitleLineClass, getTitleSizeClass } from './web-part-title-format.internal';
import type { TitleDocument, TitleInline } from './web-part-title.types.internal';
import { useWebPartTitleStyles } from './web-part-title-styles.internal';

/** Render a validated protocol run as owned React nodes; never forward arbitrary AST fields. */
function renderInline(inline: TitleInline, key: number): React.ReactNode {
  if (inline.type === 'hardBreak') return <br key={key} />;
  const marks = inline.marks;
  let content: React.ReactNode = inline.text;
  const markerClasses = [
    marks?.color && `wps-color-${marks.color}`,
    marks?.size !== undefined && getTitleSizeClass(marks.size),
  ].filter(Boolean).join(' ');
  if (markerClasses) content = <span className={markerClasses}>{content}</span>;
  if (marks?.underline) content = <u>{content}</u>;
  if (marks?.italic) content = <em>{content}</em>;
  if (marks?.bold) content = <strong>{content}</strong>;
  return <React.Fragment key={key}>{content}</React.Fragment>;
}

/** Engine-independent Read view. Invalid AST errors remain renderer errors, without persistence.
 * Empty titles return an empty Fragment, compatible with React 17 JSX mounting.
 * The shared wrapper also supplies the private editor's scoped semantic-marker styles.
 */
export function WebPartTitleReadRenderer({ document }: { document: TitleDocument }): React.ReactElement {
  const styles = useWebPartTitleStyles();
  if (!isTitleDocument(document)) throw new TypeError('Invalid title document');
  if (isTitleDocumentEmpty(document)) return <React.Fragment />;
  const blockClassName = [
    document.lineSpacing !== undefined && getTitleLineClass(document.lineSpacing),
    document.spaceAfter && 'wps-space-after',
    document.spaceBefore && 'wps-space-before',
  ].filter(Boolean).join(' ');
  const block = React.createElement(document.type === 'heading' ? `h${document.level}` : 'p',
    { className: blockClassName }, document.content.map(renderInline));
  return <div className={styles.root}>{block}</div>;
}
