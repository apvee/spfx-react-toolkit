'use strict';

const build = require('@microsoft/sp-build-web');
const path = require('path');

// The legacy Sass importer expands ~ to node_modules/ relative to the app.
// Include the resolved dependency root so hoisted workspaces and isolated consumers both work.
const dependencyRoot = path.resolve(path.dirname(require.resolve('@fluentui/react/package.json')), '../../..');
process.env.SASS_PATH = [dependencyRoot, process.env.SASS_PATH].filter(Boolean).join(path.delimiter);

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.configureWebpack.mergeConfig({
  additionalConfiguration: (configuration) => {
    const serve = require('./config/serve.json');
    const origin = `${serve.https ? 'https' : 'http'}://${serve.hostname || 'localhost'}:${serve.port || 4321}`;
    for (const plugin of configuration.plugins || []) {
      const options = plugin._options;
      // SPFx 1.21.1 omits this origin when constructing its cumulative manifest processor.
      // Keep this version-specific workaround confined to debug manifests.
      if (options && !options.production && options.cumulativeManifestProcessor) {
        options.cumulativeManifestProcessor._options.baseUrl = origin;
      }
    }
    return configuration;
  }
});

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
