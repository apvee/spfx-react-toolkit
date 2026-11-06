const path = require('node:path');
const ts = require('typescript');
const { getScenarioBlockRanges } = require('../../../scripts/docs/documentation-tsdoc.cjs');

const lineBreak = /\r\n|[\n\r\u2028\u2029]/u;

function compactComment(raw) {
  const ranges = getScenarioBlockRanges(raw);
  if (!ranges.length) return raw;

  let compact = '';
  let cursor = 0;
  for (const { pos, end } of ranges) {
    compact += raw.slice(cursor, pos);
    cursor = end;
  }
  compact += raw.slice(cursor);
  if (/^\/\*\*[\s*]*\*\/$/u.test(compact)) compact = '';

  // Move the deleted line breaks outside the comment so TypeScript emits a
  // compact docblock. Later tokens retain their original line and column.
  const originalLines = raw.split(lineBreak);
  const compactLines = compact.split(lineBreak);
  const deletedLines = originalLines.length - compactLines.length;
  const lastColumn = originalLines.at(-1).length;
  if (deletedLines) return compact + ' ' + '\n'.repeat(deletedLines) + ' '.repeat(lastColumn);
  return compact + ' '.repeat(lastColumn - compactLines.at(-1).length);
}

function stripScenarioComments(source) {
  const { text } = source;
  const comments = new Map();
  function visit(node) {
    // The compiler also emits detached and trailing doc comments. Inspect
    // actual token trivia while leaving string, template and JSX text intact.
    if (node.kind !== ts.SyntaxKind.JsxText) {
      const pos = node.getFullStart();
      const limit = node.getStart(source);
      if (pos < limit) {
        const leading = ts.getLeadingCommentRanges(text, pos) ?? [];
        const trailing = ts.getTrailingCommentRanges(text, pos) ?? [];
        for (const comment of [...leading, ...trailing]) {
          if (comment.end <= limit && text.startsWith('/**', comment.pos)) comments.set(comment.pos, comment.end);
        }
      }
    }
    for (const child of node.getChildren(source)) visit(child);
  }
  visit(source);
  let filtered = text;
  for (const [pos, end] of [...comments].sort(([left], [right]) => right - left)) {
    const raw = text.slice(pos, end);
    if (raw.includes('@scenario')) filtered = filtered.slice(0, pos) + compactComment(raw) + filtered.slice(end);
  }
  return filtered;
}

function compile(project) {
  const formatHost = {
    getCanonicalFileName: ts.createGetCanonicalFileName(ts.sys.useCaseSensitiveFileNames),
    getCurrentDirectory: ts.sys.getCurrentDirectory,
    getNewLine: () => ts.sys.newLine,
  };
  const report = diagnostics => ts.sys.write(ts.formatDiagnostics(diagnostics, formatHost));
  const config = ts.getParsedCommandLineOfConfigFile(project, {}, {
    ...ts.sys,
    onUnRecoverableConfigFileDiagnostic: diagnostic => report([diagnostic]),
  });
  if (!config) return ts.ExitStatus.DiagnosticsPresent_OutputsSkipped;
  if (config.errors.length) {
    report(config.errors);
    return ts.ExitStatus.DiagnosticsPresent_OutputsSkipped;
  }

  const host = ts.createCompilerHost(config.options);
  const root = path.resolve(config.options.rootDir ?? path.dirname(project));
  const getSourceFile = host.getSourceFile;
  host.getSourceFile = (fileName, languageVersion, onError, shouldCreateNewSourceFile) => {
    const source = getSourceFile(fileName, languageVersion, onError, shouldCreateNewSourceFile);
    const relative = path.relative(root, fileName);
    if (!source?.text.includes('@scenario') || relative === '..' || relative.startsWith('..' + path.sep) ||
      path.isAbsolute(relative) || relative.split(path.sep).includes('node_modules') || !/\.tsx?$/u.test(relative)) return source;

    const filtered = stripScenarioComments(source);
    if (filtered === source.text) return source;
    const input = ts.createSourceFile(fileName, filtered, languageVersion, false, source.scriptKind);
    // TypeScript reads file pragmas before it interprets JSDoc. Keep its
    // original compiler settings even when their spelling was in a scenario.
    input.pragmas = source.pragmas;
    languageVersion.setExternalModuleIndicator?.(input);
    return input;
  };
  const program = ts.createProgram({
    rootNames: config.fileNames,
    options: config.options,
    projectReferences: config.projectReferences,
    host,
  });
  const emit = program.emit();
  const diagnostics = ts.sortAndDeduplicateDiagnostics([...ts.getPreEmitDiagnostics(program), ...emit.diagnostics]);
  if (!diagnostics.length) return ts.ExitStatus.Success;
  report(diagnostics);
  return emit.emitSkipped ? ts.ExitStatus.DiagnosticsPresent_OutputsSkipped : ts.ExitStatus.DiagnosticsPresent_OutputsGenerated;
}

process.exitCode = compile(path.resolve(process.argv[2] ?? 'tsconfig.json'));
