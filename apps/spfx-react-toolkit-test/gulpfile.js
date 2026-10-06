'use strict';

const build = require('@microsoft/sp-build-web');
const path = require('path');

// The legacy Sass importer expands ~ to node_modules/ relative to the app.
// Include the resolved dependency root so hoisted workspaces and isolated consumers both work.
const dependencyRoot = path.resolve(path.dirname(require.resolve('@fluentui/react/package.json')), '../../..');
process.env.SASS_PATH = [dependencyRoot, process.env.SASS_PATH].filter(Boolean).join(path.delimiter);

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
