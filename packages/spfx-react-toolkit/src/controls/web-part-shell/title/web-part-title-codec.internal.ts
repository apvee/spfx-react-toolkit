import type { TitleDocument, TitleInline, TitleMarks, TitleParseResult } from './web-part-title.types.internal';
import {
  getTitleClassColor, getTitleClassLine, getTitleClassSize, getTitleLineClass,
  isTitleColor, isTitleHeadingLevel, isTitleLineSpacing, isTitleSize, titleFormatRegistry
} from './web-part-title-format.internal';

// These subtrees contain executable, embedded or non-title content, rather than harmless wrappers.
const discardedTags = new Set([
  'script', 'style', 'iframe', 'object', 'embed', 'svg', 'math', 'template', 'noscript',
  'form', 'input', 'button', 'select', 'textarea', 'meta', 'link', 'base',
  'audio', 'video', 'img', 'source', 'track'
]);
const boundaryTags = new Set([
  'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'div', 'section', 'article', 'header',
  'footer', 'blockquote', 'ul', 'ol', 'li', 'table', 'tr', 'td', 'th', 'pre', 'hr'
]);

function emptyDocument(): TitleDocument {
  return { type: 'heading', level: 2, content: [] };
}

function literalDocument(title: string): TitleDocument {
  return { type: 'heading', level: 2, content: title ? [{ type: 'text', text: title }] : [] };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  // DOM/JSON callers may supply legacy null; it is rejected, never used as an absence value.
  return typeof value === 'object' && !!value && !Array.isArray(value);
}

function isMarks(value: unknown): value is TitleMarks {
  return isRecord(value) &&
    (value.bold === undefined || value.bold === true) &&
    (value.italic === undefined || value.italic === true) &&
    (value.underline === undefined || value.underline === true) &&
    (value.size === undefined || isTitleSize(value.size)) &&
    (value.color === undefined || isTitleColor(value.color));
}

function isInline(value: unknown): value is TitleInline {
  if (!isRecord(value)) return false;
  return value.type === 'hardBreak' || (value.type === 'text' && typeof value.text === 'string' &&
    (value.marks === undefined || isMarks(value.marks)));
}

/** Validate unknown engine/JSON boundaries before consuming a document. */
export function isTitleDocument(value: unknown): value is TitleDocument {
  if (!isRecord(value)) return false;
  return (value.type === 'paragraph' || (value.type === 'heading' && isTitleHeadingLevel(value.level))) &&
    Array.isArray(value.content) && Array.from(value.content).every(isInline) &&
    (value.lineSpacing === undefined || isTitleLineSpacing(value.lineSpacing)) &&
    (value.spaceBefore === undefined || value.spaceBefore === true) &&
    (value.spaceAfter === undefined || value.spaceAfter === true);
}

function validateDocument(value: unknown): asserts value is TitleDocument {
  if (!isTitleDocument(value)) throw new TypeError('Invalid title document');
}

function sameMarks(left: TitleMarks | undefined, right: TitleMarks | undefined): boolean {
  return left?.bold === right?.bold && left?.italic === right?.italic &&
    left?.underline === right?.underline && left?.size === right?.size && left?.color === right?.color;
}

function appendInline(target: TitleInline[], inline: TitleInline): void {
  if (inline.type === 'hardBreak') {
    target.push({ type: 'hardBreak' });
    return;
  }
  if (!inline.text) return;
  const previous = target[target.length - 1];
  if (previous?.type === 'text' && sameMarks(previous.marks, inline.marks)) previous.text += inline.text;
  else target.push({ ...inline });
}

function readMarks(element: Element, inherited: TitleMarks): TitleMarks {
  const marks = { ...inherited };
  const tag = element.localName;
  if (tag === 'strong' || tag === 'b') marks.bold = true;
  if (tag === 'em' || tag === 'i') marks.italic = true;
  if (tag === 'u') marks.underline = true;
  for (const className of Array.from(element.classList)) {
    const size = getTitleClassSize(className);
    const color = getTitleClassColor(className);
    if (size !== undefined) marks.size = size;
    if (color !== undefined) marks.color = color;
  }
  return marks;
}

function readBlock(element: Element): TitleDocument | undefined {
  const tag = element.localName;
  const level = /^h[1-4]$/.test(tag) ? Number(tag.slice(1)) : undefined;
  const block: TitleDocument | undefined = tag === 'p'
    ? { type: 'paragraph', content: [] }
    : isTitleHeadingLevel(level) ? { type: 'heading', level, content: [] } : undefined;
  if (!block) return undefined;
  for (const className of Array.from(element.classList)) {
    const line = getTitleClassLine(className);
    if (line !== undefined) block.lineSpacing = line;
    if (className === 'wps-space-before') block.spaceBefore = true;
    if (className === 'wps-space-after') block.spaceAfter = true;
  }
  // Read only the finite native line-height declaration. All other CSS is discarded.
  // getAttribute uses the legacy DOM null sentinel; convert immediately to undefined.
  const style = element.getAttribute('style') ?? undefined;
  if (style && block.lineSpacing === undefined) {
    for (const declaration of style.split(';')) {
      const match = /^\s*line-height\s*:\s*(1|1\.2|1\.4|1\.6|2|2\.4)\s*$/i.exec(declaration);
      const value = match ? Number(match[1]) : undefined;
      if (isTitleLineSpacing(value)) block.lineSpacing = value;
    }
  }
  return block;
}

function withinDepthLimit(root: DocumentFragment): boolean {
  const pending: { node: Node; depth: number }[] = Array.from(root.childNodes).map(node => ({ node, depth: 1 }));
  while (pending.length) {
    const current = pending.pop();
    if (!current) break;
    if (current.node.nodeType !== 1) continue;
    if (current.depth > titleFormatRegistry.depthLimit) return false;
    for (const child of Array.from(current.node.childNodes)) pending.push({ node: child, depth: current.depth + 1 });
  }
  return true;
}

/** Parse in the provided owner's inert template; original HTML is never attached to visible DOM. */
export function parseTitleHTML(html: string, ownerDocument: Document): TitleParseResult {
  if (typeof html !== 'string' || new TextEncoder().encode(html).byteLength > titleFormatRegistry.htmlByteLimit) {
    return { kind: 'invalid', document: emptyDocument() };
  }
  const template = ownerDocument.createElement('template');
  template.innerHTML = html;
  if (!withinDepthLimit(template.content)) return { kind: 'invalid', document: emptyDocument() };
  let firstBlock: TitleDocument | undefined;
  interface InlineSegment {
    content: TitleInline[];
    startsWithBlock: boolean;
    endsWithBlock: boolean;
    /** Empty blocks still participate in paste boundaries. */
    hasSegment: boolean;
  }
  function children(parent: Node, inherited: TitleMarks): InlineSegment {
    const segment: InlineSegment = { content: [], startsWithBlock: false, endsWithBlock: false, hasSegment: false };
    function append(inner: InlineSegment): void {
      if (!inner.hasSegment) return;
      if (!segment.hasSegment) segment.startsWithBlock = inner.startsWithBlock;
      else if (segment.endsWithBlock || inner.startsWithBlock) segment.content.push({ type: 'hardBreak' });
      for (const inline of inner.content) appendInline(segment.content, inline);
      segment.endsWithBlock = inner.endsWithBlock;
      segment.hasSegment = true;
    }
    for (const node of Array.from(parent.childNodes)) {
      if (node.nodeType === 3) {
        const text = node.textContent ?? '';
        // Formatting whitespace between pasted blocks is not another text run.
        if (!text.trim() && (segment.endsWithBlock || (parent.nodeType !== 1 && !segment.hasSegment))) continue;
        const marks = Object.keys(inherited).length ? inherited : undefined;
        append({
          content: [marks ? { type: 'text', text, marks } : { type: 'text', text }],
          startsWithBlock: false, endsWithBlock: false, hasSegment: !!text
        });
        continue;
      }
      if (node.nodeType !== 1) continue;
      // nodeType establishes Element across owner-document/iframe realms without global instanceof.
      const element = node as Element;
      if (discardedTags.has(element.localName)) continue;
      if (element.localName === 'br') {
        append({ content: [{ type: 'hardBreak' }], startsWithBlock: false, endsWithBlock: false, hasSegment: true });
        continue;
      }
      const block = readBlock(element);
      if (!firstBlock && block) firstBlock = block;
      const isBlock = boundaryTags.has(element.localName);
      const inner = children(element, readMarks(element, inherited));
      append({
        content: inner.content,
        startsWithBlock: isBlock || inner.startsWithBlock,
        endsWithBlock: isBlock || inner.endsWithBlock,
        hasSegment: isBlock || inner.hasSegment
      });
    }
    return segment;
  }
  const content = children(template.content, {}).content;
  const document: TitleDocument = firstBlock || emptyDocument();
  document.content = content;
  try {
    const canonical = serializeTitleDocument(document);
    return { kind: canonical === html ? 'valid' : 'normalized', document };
  } catch (error) {
    // A readable source can expand beyond the protocol budget after escaping/normalization.
    if (error instanceof RangeError) return { kind: 'invalid', document: emptyDocument() };
    throw error;
  }
}

/** Undefined HTML uses literal title; defined empty HTML is authoritative. No property writes. */
export function resolveWebPartTitle(title: string, titleHTML: string | undefined, ownerDocument: Document): TitleParseResult {
  if (titleHTML === undefined) return { kind: 'valid', document: literalDocument(title) };
  const result = parseTitleHTML(titleHTML, ownerDocument);
  return result.kind === 'invalid' ? { kind: 'invalid', document: literalDocument(title) } : result;
}

/** Search text derives exclusively from the validated AST with whitespace and break normalization. */
export function getTitlePlainText(document: TitleDocument): string {
  validateDocument(document);
  return document.content.map(inline => inline.type === 'text' ? inline.text : ' ').join('').replace(/\s+/g, ' ').trim();
}

/** Whitespace and hard breaks alone do not render a title or persist an empty heading. */
export function isTitleDocumentEmpty(document: TitleDocument): boolean {
  return getTitlePlainText(document) === '';
}

function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function serializeInline(inline: TitleInline): string {
  if (inline.type === 'hardBreak') return '<br>';
  const marks = inline.marks;
  let html = escapeText(inline.text);
  const classes: string[] = [];
  if (marks?.color !== undefined) classes.push(`wps-color-${marks.color}`);
  if (marks?.size !== undefined) classes.push(`wps-size-${marks.size}`);
  if (classes.length) html = `<span class="${classes.sort().join(' ')}">${html}</span>`;
  if (marks?.underline) html = `<u>${html}</u>`;
  if (marks?.italic) html = `<em>${html}</em>`;
  if (marks?.bold) html = `<strong>${html}</strong>`;
  return html;
}

/**
 * Deterministic, escaped single-block protocol HTML; empty documents serialize as an empty string.
 * @throws TypeError for malformed AST input.
 * @throws RangeError when canonical HTML exceeds 65,536 UTF-8 bytes. Editor callers must retain
 * their current document, show a recoverable size error and skip the property patch; never truncate.
 */
export function serializeTitleDocument(document: TitleDocument): string {
  validateDocument(document);
  if (isTitleDocumentEmpty(document)) return '';
  const tag = document.type === 'heading' ? `h${document.level}` : 'p';
  const classes: string[] = [];
  if (document.lineSpacing !== undefined) classes.push(getTitleLineClass(document.lineSpacing));
  if (document.spaceBefore) classes.push('wps-space-before');
  if (document.spaceAfter) classes.push('wps-space-after');
  const attr = classes.length ? ` class="${classes.sort().join(' ')}"` : '';
  const content: TitleInline[] = [];
  for (const inline of document.content) appendInline(content, inline);
  const html = `<${tag}${attr}>${content.map(serializeInline).join('')}</${tag}>`;
  if (new TextEncoder().encode(html).byteLength > titleFormatRegistry.htmlByteLimit) {
    throw new RangeError('Title HTML exceeds 65536 UTF-8 bytes');
  }
  return html;
}
