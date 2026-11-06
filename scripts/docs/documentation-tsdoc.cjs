const { TSDocConfiguration, TSDocParser, TSDocTagDefinition, TSDocTagSyntaxKind } = require('@microsoft/tsdoc');
const MarkdownIt = require('markdown-it');
const profile = require('../../packages/spfx-react-toolkit/tsdoc.json');

const markdown = new MarkdownIt({ html: false, linkify: false, typographer: false });
const parserConfig = new TSDocConfiguration();
for (const tag of profile.tagDefinitions) parserConfig.addTagDefinition(new TSDocTagDefinition({ ...tag, syntaxKind: TSDocTagSyntaxKind.BlockTag }));
const parser = new TSDocParser(parserConfig);

function excerpts(node) {
  if (node.kind === 'Excerpt') return [node];
  return node.getChildNodes().flatMap(excerpts);
}
function nodeRange(node, raw) {
  const ranges = excerpts(node).filter(item => !item.content.isEmpty()).map(item => item.content.getContainingTextRange());
  if (!ranges.length) return undefined;
  const pos = Math.min(...ranges.map(item => item.pos));
  const end = Math.max(...ranges.map(item => item.end));
  // The parser's adapted delimiter has three backticks followed by spaces.
  // Restore any original trailing delimiter characters at the excerpt boundary.
  const delimiter = raw?.slice(end - 3, end);
  const tail = delimiter === '```' ? /^`+/.exec(raw.slice(end)) : delimiter === '~~~' ? /^~+/.exec(raw.slice(end)) : undefined;
  return { pos, end: end + (tail?.[0].length ?? 0) };
}

/** Adapt Markdown fences to the pinned TSDoc grammar without changing offsets. */
function parserText(raw) {
  const original = raw.replace(/\r(?!\n)|[\u2028\u2029]/gu, '\n').split('\n');
  const body = original.map((line, index) => line.replace(index === 0 ? /^\s*\/\*\* ?/ : /^\s*\* ?/, '').replace(/\*\/\s*$/, ''));
  const adapted = [...original];
  for (const token of markdown.parse(body.join('\n'), {})) {
    if (token.type !== 'fence') continue;
    const [start, end] = token.map;
    const rewriteDelimiter = line => line.replace(/[`~]{3,}/, run => '```' + ' '.repeat(run.length - 3));
    adapted[start] = rewriteDelimiter(adapted[start]);
    const last = end - 1;
    const hasClose = last > start && /^\s*([`~])\1{2,}\s*$/.test(body[last]);
    if (hasClose) adapted[last] = rewriteDelimiter(adapted[last]);
    for (let index = start + 1; index < (hasClose ? last : end); index += 1) {
      // An inner shorter backtick run is code in Markdown, but would close a
      // three-backtick TSDoc fence. Its original spelling stays in the model.
      adapted[index] = adapted[index].replace(/^(\s*\*?\s*)`{3}/, '$1~``');
    }
  }
  return adapted.join('\n');
}

function parseDocumentationComment(raw) {
  return parser.parseString(parserText(raw));
}

const lineBreak = /\r\n|[\n\r\u2028\u2029]/u;
const docLinePrefix = /^[\t ]*(?:\*[\t ]*)?$/u;
const linePrefix = (raw, pos) => raw.slice(0, pos).split(lineBreak).at(-1);

function wholeDocLines(raw, { pos, end }) {
  const prefix = linePrefix(raw, pos);
  if (!docLinePrefix.test(prefix)) return { pos, end };
  const lineEnd = raw.slice(end).match(/^[\t ]*(?:\r\n|[\n\r\u2028\u2029])/u);
  return {
    pos: pos - prefix.length,
    end: end + (lineEnd?.[0].length ?? 0),
  };
}

/** Return source-relative removal ranges, preserving comment delimiters and sibling tags. */
function getScenarioBlockRanges(raw) {
  const { docComment, log } = parseDocumentationComment(raw);
  const scenarios = docComment.customBlocks.filter(block => block.blockTag.tagName === '@scenario')
    .map(block => nodeRange(block, raw));
  // TSDoc stores modifiers separately but can include their spelling in the
  // preceding block's excerpts. Unknown standalone tags can also be compiler
  // directives. Preserve both outside fences, which the parser treats as code.
  const modifiers = docComment.modifierTagSet.nodes.map(node => nodeRange(node, raw));
  const unknownTags = log.messages.flatMap(message => {
    const pos = message.textRange?.pos;
    if (pos === undefined || !['tsdoc-undefined-tag', 'tsdoc-unsupported-tag', 'tsdoc-characters-after-block-tag'].includes(message.messageId) ||
      !/^@(?!scenario\b)[a-z][\w-]*/iu.test(raw.slice(pos)) || !docLinePrefix.test(linePrefix(raw, pos))) return [];
    const line = raw.slice(pos).split(lineBreak)[0];
    const lineEnd = pos + (line.includes('*/') ? line.indexOf('*/') : line.length);
    const end = scenarios.find(scenario => pos < scenario.pos && scenario.pos < lineEnd)?.pos ?? lineEnd;
    return [{ pos, end }];
  });
  const preserved = [];
  for (const tag of [...modifiers, ...unknownTags].map(range => wholeDocLines(raw, range)).sort((a, b) => a.pos - b.pos)) {
    const previous = preserved.at(-1);
    if (previous && (tag.pos <= previous.end || /^[\t ]*$/u.test(raw.slice(previous.end, tag.pos)))) {
      previous.end = Math.max(previous.end, tag.end);
    } else preserved.push(tag);
  }
  return scenarios
    .flatMap(scenario => {
      const { pos, end } = wholeDocLines(raw, scenario);
      const removals = [];
      let cursor = pos;
      for (const tag of preserved) {
        if (tag.end <= cursor || tag.pos >= end) continue;
        if (cursor < tag.pos) removals.push({ pos: cursor, end: tag.pos });
        cursor = tag.end;
      }
      if (cursor < end) removals.push({ pos: cursor, end });
      return removals;
    })
    .sort((a, b) => a.pos - b.pos);
}

module.exports = { parseDocumentationComment, nodeRange, getScenarioBlockRanges };
