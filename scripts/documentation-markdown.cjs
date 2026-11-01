const path = require('node:path');
const MarkdownIt = require('markdown-it');
// Unmodified read-only parser configuration is reusable; each parse keeps its
// tokens/environment local. The source transform still owns an isolated parser.
const markdownReader = new MarkdownIt({ html: true });

/** Slug rendered heading text or a literal symbol name; no Markdown parsing. */
function markdownHeadingSlug(heading) {
  return heading.toLowerCase()
    .replace(/[^\p{L}\p{N}_ -]/gu, '').replace(/ /g, '-');
}
const externalTarget = target => /^(?:[a-z][a-z0-9+.-]*:|\/\/)/i.test(target);
const editorialRoute = route => route.replace(/^docs\/_sources\/(?=INDEX\.md$|getting-started\/|guides\/|development\/)/, 'docs/');
function relativeMarkdownTarget(from, route, suffix = '') {
  return (path.posix.relative(path.posix.dirname(from), route) || path.posix.basename(route)) + suffix;
}
function rebaseMarkdownTarget(target, source, destination, publicOrigin) {
  if (!target || externalTarget(target) || target.startsWith('/')) return target;
  if (target.startsWith('#')) {
    const origin = publicOrigin || (source.endsWith('.md') ? editorialRoute(source) : undefined);
    if (!origin) throw new Error(`Fragment link requires a public page origin for ${source}.`);
    return origin === destination ? target : relativeMarkdownTarget(destination, origin, target);
  }
  const split = target.search(/[?#]/);
  const file = split < 0 ? target : target.slice(0, split);
  const suffix = split < 0 ? '' : target.slice(split);
  return relativeMarkdownTarget(destination, editorialRoute(path.posix.normalize(path.posix.join(path.posix.dirname(source), file))), suffix);
}
const escapeLabel = value => value.replace(/\\/g, '\\\\').replace(/[\[\]]/g, '\\$&');
const destinationText = (target, angled) => target.replace(/[<>\s]/g, character => encodeURIComponent(character))
  .replace(angled ? /\\/g : /[\\()]/g, '\\$&');

/**
 * Source-preserving adapter for pinned markdown-it 14.1.0. Only destinations,
 * source references and explicit directive spans are patched. Code, Markdown
 * layout, titles, escapes and ordinary HTML comments retain their source bytes.
 * The parser owns syntax recognition (including reference definitions); this is
 * deliberately not a second Markdown/JSDoc parser.
 * @param {string} source Markdown with LF line endings.
 * @param {{link?: (target: string) => string, reference?: (reference: string, label?: string) => string, include?: (directive: string) => string}} [handlers]
 * @returns {string} Markdown with selected source spans replaced.
 */
function transformMarkdown(source, handlers = {}) {
  const md = new MarkdownIt({ html: true });
  // markdown-it's helpers are shared between instances. Never mutate the
  // exported singleton when instrumenting one parse.
  md.helpers = { ...md.helpers };
  const edits = new Map();
  const add = (start, end, value) => edits.set(start + ':' + end, { start, end, value });
  const offsets = [0];
  for (let i = 0; i < source.length; i++) if (source[i] === '\n') offsets.push(i + 1);
  function sourceMap(text, start, end) {
    const result = [];
    let cursor = start;
    // Inline content differs only by block prefixes/indentation stripped by
    // markdown-it. Match each line within its original physical line first;
    // a character-level subsequence would mis-map list/blockquote markers.
    for (const [index, line] of text.split('\n').entries()) {
      if (index) { result.push(cursor - 1); }
      const lineEnd = source.indexOf('\n', cursor);
      const limit = lineEnd < 0 ? end : Math.min(end, lineEnd);
      const at = source.indexOf(line, cursor);
      if (at >= cursor && at + line.length <= limit) {
        for (let i = 0; i < line.length; i++) result.push(at + i);
      } else {
        // The table rule removes the escape before a cell's literal pipe.
        // Account for that one pinned-parser normalization, preserving bytes.
        let mapping;
        for (let candidate = cursor; candidate < limit && !mapping; candidate++) {
          const attempt = [];
          let position = candidate;
          for (const character of line) {
            if (character === '|' && source[position] === '\\' && source[position + 1] === '|') position++;
            if (source[position] !== character || position >= limit) break;
            attempt.push(position++);
          }
          if (attempt.length === line.length) mapping = attempt;
        }
        if (!mapping) throw new Error('Markdown source mapping failed.');
        result.push(...mapping);
      }
      cursor = limit + 1;
    }
    result.push(result.length ? result.at(-1) + 1 : start);
    return result;
  }
  let inlineMap;
  function destination(record, map) {
    if (!handlers.link || !record.result.ok) return;
    const { str, start, result } = record;
    const angled = str[start] === '<';
    const from = start + (angled ? 1 : 0);
    const to = result.pos - (angled ? 1 : 0);
    const target = handlers.link(result.str);
    if (target !== result.str) add(map[from], map[to - 1] + 1, destinationText(target, angled));
  }
  for (const name of ['link', 'image']) {
    const original = md.inline.ruler.__rules__.find(rule => rule.name === name).fn;
    md.inline.ruler.at(name, (state, silent) => {
      const calls = [];
      const parse = md.helpers.parseLinkDestination;
      md.helpers.parseLinkDestination = (str, start, max) => {
        const result = parse(str, start, max);
        if (str === state.src) calls.push({ str, start, result });
        return result;
      };
      const parseInline = md.inline.parse;
      if (name === 'image') {
        const labelStart = state.pos + 2;
        const parentMap = inlineMap;
        md.inline.parse = function (text, parser, env, tokens) {
          const at = state.src.indexOf(text, labelStart);
          inlineMap = parentMap?.slice(at, at + text.length + 1);
          try { return parseInline.call(this, text, parser, env, tokens); }
          finally { inlineMap = parentMap; }
        };
      }
      let accepted;
      try { accepted = original(state, silent); }
      finally { md.helpers.parseLinkDestination = parse; md.inline.parse = parseInline; }
      if (accepted && !silent && inlineMap) for (const call of calls) destination(call, inlineMap);
      return accepted;
    });
  }
  const referenceRule = md.block.ruler.__rules__.find(rule => rule.name === 'reference').fn;
  md.block.ruler.at('reference', (state, startLine, endLine, silent) => {
    const calls = [];
    const parse = md.helpers.parseLinkDestination;
    md.helpers.parseLinkDestination = (str, start, max) => {
      const result = parse(str, start, max);
      calls.push({ str, start, result });
      return result;
    };
    let accepted;
    try { accepted = referenceRule(state, startLine, endLine, silent); }
    finally { md.helpers.parseLinkDestination = parse; }
    if (accepted && !silent) for (const call of calls) {
      const start = state.bMarks[startLine] + state.tShift[startLine];
      destination(call, sourceMap(call.str.replace(/\n$/, ''), start, offsets[state.line] ?? source.length));
    }
    return accepted;
  });
  function directives(text, map) {
    for (const match of text.matchAll(/<!--\s*toolkit:include\b[\s\S]*?(?:-->|$)/g)) {
      if (!handlers.include) throw new Error('Unresolved or recursive toolkit:include directive.');
      add(map[match.index], map[match.index + match[0].length - 1] + 1, handlers.include(match[0]));
    }
  }
  md.inline.ruler.before('text', 'documentation_reference', (state, silent) => {
    if (!state.src.startsWith('{@link', state.pos)) return false;
    const match = state.src.slice(state.pos).match(/^\{@link\s+([^}\n]+)\}/);
    if (!match) throw new Error('Invalid documentation reference syntax.');
    if (!silent) {
      if (!handlers.reference) throw new Error('Unresolved documentation reference.');
      const [reference, ...label] = match[1].split('|');
      const value = handlers.reference(reference.trim(), label.length ? label.join('|').trim() : undefined);
      add(inlineMap[state.pos], inlineMap[state.pos + match[0].length - 1] + 1, value);
      state.pending += match[0];
    }
    state.pos += match[0].length;
    return true;
  });
  const htmlRule = md.inline.ruler.__rules__.find(rule => rule.name === 'html_inline').fn;
  md.inline.ruler.at('html_inline', (state, silent) => {
    const start = state.pos;
    const result = htmlRule(state, silent);
    if (result && !silent) directives(state.src.slice(start, state.pos), inlineMap.slice(start, state.pos + 1));
    return result;
  });
  md.core.ruler.at('inline', state => {
    let cursor = 0;
    for (const token of state.tokens) {
      if (token.map) cursor = offsets[token.map[0]];
      if (token.type === 'inline') {
        inlineMap = sourceMap(token.content, cursor, token.map ? offsets[token.map[1]] ?? source.length : source.length);
        cursor = inlineMap.at(-1);
        md.inline.parse(token.content, md, state.env, token.children);
      } else if (token.type === 'html_block') {
        const start = offsets[token.map[0]];
        const map = sourceMap(token.content, start, offsets[token.map[1]] ?? source.length);
        directives(token.content, map);
      }
    }
  });
  md.parse(source, {});
  const ordered = [...edits.values()].sort((a, b) => b.start - a.start);
  let previous = source.length;
  for (const edit of ordered) {
    if (edit.end > previous) throw new Error('Overlapping Markdown transformations.');
    source = source.slice(0, edit.start) + edit.value + source.slice(edit.end);
    previous = edit.start;
  }
  return source;
}

/**
 * Resolve normalized prose using filesystem provenance for relative file links
 * and publicOrigin for fragment-only links. Markdown authoring paths infer their
 * public route; TypeScript contracts require a known canonical public page.
 * A fragment stays local only when that origin is also the destination.
 */
function resolveMarkdown(source, sourceFile, destination, resolveReference, include, publicOrigin) {
  return transformMarkdown(source, {
    link: target => rebaseMarkdownTarget(target, sourceFile, destination, publicOrigin),
    reference: (reference, label) => {
      if (externalTarget(reference)) return `[${escapeLabel(label || reference)}](${destinationText(reference, false)})`;
      const resolved = resolveReference(reference);
      if (!resolved.page) throw new Error(`Documentation reference ${reference} has no page.`);
      return `[${escapeLabel(label || reference)}](${relativeMarkdownTarget(destination, resolved.page.route, '#' + resolved.anchor)})`;
    },
    include,
  });
}

/** Rendered inline text excludes markup/destinations and retains literal code. */
function markdownInlineText(tokens) {
  return tokens.map(token => token.type === 'text' || token.type === 'code_inline' ? token.content
    : token.type === 'softbreak' || token.type === 'hardbreak' ? ' '
      : token.children ? markdownInlineText(token.children) : '').join('');
}
/** Safe plain visible label for a link containing already resolved Markdown. */
function markdownLabel(source) {
  return escapeLabel(markdownInlineText(markdownReader.parseInline(source, {}))).replace(/[`*_]/g, '\\$&');
}
/** Natural anchor for a rendered inline heading (symbol slugs stay plain text). */
function markdownHeadingAnchor(heading) {
  const tokens = markdownReader.parseInline(heading, {});
  return markdownHeadingSlug(markdownInlineText(tokens));
}
/** Extract natural duplicate-aware anchors and explicit HTML IDs, excluding code. */
function markdownAnchors(source) {
  const tokens = markdownReader.parse(source, {});
  const used = new Set();
  const anchors = [];
  function explicit(tokens) {
    for (const token of tokens) {
      if (token.type === 'html_inline' || token.type === 'html_block') {
        for (const match of token.content.matchAll(/<a\s+id=["']([^"']+)["']/g)) if (!anchors.includes(match[1])) anchors.push(match[1]);
      }
      if (token.children) explicit(token.children);
    }
  }
  explicit(tokens);
  for (let i = 0; i < tokens.length; i++) if (tokens[i].type === 'heading_open') {
    const base = markdownHeadingSlug(markdownInlineText(tokens[i + 1].children));
    let value = base;
    let suffix = 0;
    while (used.has(value)) value = base + '-' + (++suffix);
    used.add(value);
    anchors.push(value);
  }
  return anchors;
}
/** Collect parser-recognized links/images (including reference-style usages). */
function markdownLinks(source) {
  const links = [];
  function visit(tokens) {
    for (const token of tokens) {
      const target = token.type === 'link_open' ? token.attrGet('href') : token.type === 'image' ? token.attrGet('src') : undefined;
      if (target !== undefined) links.push(target);
      if (token.children) visit(token.children);
    }
  }
  visit(markdownReader.parse(source, {}));
  return links;
}
module.exports = { markdownLabel, markdownHeadingSlug, markdownHeadingAnchor, markdownAnchors, markdownLinks, transformMarkdown, resolveMarkdown,
  rebaseMarkdownTarget, relativeMarkdownTarget, editorialRoute, externalTarget };
