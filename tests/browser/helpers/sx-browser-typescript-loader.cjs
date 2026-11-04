// The fixture uses the repository's installed TypeScript; no loader dependency.
const ts = require('typescript');
module.exports = function compile(source) {
  const result = ts.transpileModule(source, { fileName: this.resourcePath,
    compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext, jsx: ts.JsxEmit.React, sourceMap: true },
    reportDiagnostics: true });
  for (const diagnostic of result.diagnostics || []) {
    if (diagnostic.category === ts.DiagnosticCategory.Error) this.emitError(new Error(ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n')));
  }
  this.callback(null, result.outputText, result.sourceMapText && JSON.parse(result.sourceMapText));
};
