import * as React from 'react';
import { Link } from '@fluentui/react-components';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';

const processor = unified().use(remarkParse).use(remarkGfm);
const maximumBytes = 64 * 1024;
const maximumNodes = 2000;
const maximumDepth = 32;

interface NodePosition {
  readonly position: { readonly start: { readonly offset: number }; readonly end: { readonly offset: number } };
}
type InlineNode = NodePosition & (
  | { readonly type: 'text' | 'inlineCode' | 'html'; readonly value: string }
  | { readonly type: 'break' }
  | { readonly type: 'emphasis' | 'strong' | 'delete'; readonly children: readonly InlineNode[] }
  // Remark's native link title uses null when no title was authored.
  // eslint-disable-next-line @rushstack/no-new-null
  | { readonly type: 'link'; readonly url: string; readonly title?: string | null; readonly children: readonly InlineNode[] });
interface Paragraph {
  readonly type: 'paragraph';
  readonly children: readonly InlineNode[];
  readonly position: { readonly start: { readonly offset: number }; readonly end: { readonly offset: number } };
}
type ParsedMarkdown =
  | { readonly kind: 'parsed'; readonly leadingWhitespace: string; readonly trailingWhitespace: string; readonly children: readonly InlineNode[] }
  | { readonly kind: 'literal'; readonly source: string; readonly reason: 'unsupported' | 'limit' | 'parse' };
interface ContentOptions {
  readonly markdown: boolean;
  readonly codeClassName: string;
  readonly linkClassName: string;
}
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}
function hasChildren(value: Record<string, unknown>): value is Record<string, unknown> & { children: unknown[] } {
  return Array.isArray(value.children);
}
function isInlineNode(value: unknown): value is InlineNode {
  if (!isRecord(value) || !hasNodePosition(value)) return false;
  switch (value.type) {
    case 'text': case 'html': case 'inlineCode':
      return typeof value.value === 'string' && value.children === undefined;
    case 'break': return value.children === undefined;
    case 'emphasis': case 'strong': case 'delete': return hasChildren(value);
    case 'link': return hasChildren(value) && typeof value.url === 'string'
      && (value.title === undefined || value.title === null || typeof value.title === 'string');
    default: return false;
  }
}
function hasNodePosition(value: Record<string, unknown>): value is Record<string, unknown> & NodePosition {
  const position = value.position;
  return isRecord(position) && isRecord(position.start) && isRecord(position.end)
    && Number.isInteger(position.start.offset) && Number.isInteger(position.end.offset);
}
/** Validate the original tree iteratively before any recursive rendering. */
function validateTree(tree: unknown): 'valid' | 'unsupported' | 'limit' {
  const pending: { node: unknown; depth: number }[] = [{ node: tree, depth: 1 }];
  let count = 0;
  let supported = true;
  while (pending.length) {
    const next = pending.pop();
    if (!next) break;
    if (++count > maximumNodes || next.depth > maximumDepth) return 'limit';
    if (!isRecord(next.node) || typeof next.node.type !== 'string') { supported = false; continue; }
    if (next.depth === 1) {
      if (next.node.type !== 'root' || !hasChildren(next.node) || next.node.children.length !== 1) supported = false;
    } else if (next.depth === 2) {
      if (next.node.type !== 'paragraph' || !hasChildren(next.node) || !hasNodePosition(next.node)) supported = false;
    } else if (!isInlineNode(next.node)) supported = false;
    if (hasChildren(next.node)) {
      for (const child of next.node.children) pending.push({ node: child, depth: next.depth + 1 });
    } else if ('children' in next.node && next.node.children !== undefined) supported = false;
  }
  return supported ? 'valid' : 'unsupported';
}

/** Parse the supported single-paragraph subset, retaining native Remark nodes. */
export function parseInlineMarkdown(source: string): ParsedMarkdown {
  if (new TextEncoder().encode(source).byteLength > maximumBytes) return { kind: 'literal', source, reason: 'limit' };
  if (/^\s*$/.test(source)) return { kind: 'parsed', leadingWhitespace: source, trailingWhitespace: '', children: [] };
  try {
    const tree = processor.parse(source);
    const validation = validateTree(tree);
    if (validation !== 'valid') return { kind: 'literal', source, reason: validation };
    // The shape and the entire inline subtree were checked above; retain their original identity.
    const paragraph = tree.children[0] as Paragraph;
    const start = paragraph.position.start.offset;
    const last = paragraph.children[paragraph.children.length - 1];
    const end = last ? last.position.end.offset : paragraph.position.end.offset;
    if (start < 0 || end < start || end > source.length) return { kind: 'literal', source, reason: 'parse' };
    return { kind: 'parsed', leadingWhitespace: source.slice(0, start), trailingWhitespace: source.slice(end), children: paragraph.children };
  } catch {
    return { kind: 'literal', source, reason: 'parse' };
  }
}

/** Apply the URL policy to Remark's already-decoded destination. */
export function isSafeMarkdownUrl(value: string | undefined): boolean {
  if (typeof value !== 'string') return false;
  const url = value.trim();
  if (!url || url.startsWith('//') || url.includes('\\')
    || Array.from(url).some(character => character.charCodeAt(0) <= 0x20 || character.charCodeAt(0) === 0x7f)) return false;
  const scheme = /^([a-z][a-z\d+.-]*):/i.exec(url);
  if (scheme && !/^(https?|mailto|tel)$/i.test(scheme[1])) return false;
  try {
    const parsed = new URL(url, 'https://formatted-text.invalid/');
    return !scheme || /^(https?:|mailto:|tel:)$/.test(parsed.protocol);
  } catch { return false; }
}
function renderNodes(nodes: readonly InlineNode[], options: ContentOptions): React.ReactNode {
  return nodes.map((node, key) => {
    switch (node.type) {
      case 'text': case 'html': return node.value;
      case 'inlineCode': return <code key={key} className={options.codeClassName}>{node.value}</code>;
      case 'break': return <br key={key} />;
      case 'emphasis': return <em key={key}>{renderNodes(node.children, options)}</em>;
      case 'strong': return <strong key={key}>{renderNodes(node.children, options)}</strong>;
      case 'delete': return <del key={key}>{renderNodes(node.children, options)}</del>;
      case 'link': {
        const children = renderNodes(node.children, options);
        return isSafeMarkdownUrl(node.url)
          ? <Link key={key} inline href={node.url.trim()} title={node.title ?? undefined} className={options.linkClassName}>{children}</Link>
          : <React.Fragment key={key}>{children}</React.Fragment>;
      }
    }
  });
}
/** Parse strings independently and traverse only arrays and React fragments. */
export function renderFormattedTextChildren(children: React.ReactNode, options: ContentOptions): React.ReactNode {
  if (!options.markdown) return children;
  if (typeof children === 'string') {
    const parsed = parseInlineMarkdown(children);
    return parsed.kind === 'literal' ? parsed.source
      : [parsed.leadingWhitespace, renderNodes(parsed.children, options), parsed.trailingWhitespace];
  }
  if (Array.isArray(children)) {
    return children.map((child, index) => (
      <React.Fragment key={React.isValidElement(child) && child.key !== null ? `element:${child.key}` : `index:${index}`}>
        {renderFormattedTextChildren(child, options)}
      </React.Fragment>
    ));
  }
  if (React.isValidElement<{ children?: React.ReactNode }>(children) && children.type === React.Fragment) {
    return React.cloneElement(children, undefined, renderFormattedTextChildren(children.props.children, options));
  }
  return children;
}
