// NODE_OPTIONS entry point for verification runs. Inherited compiler/minifier
// workers install exact advisory routing without importing metadata eagerly.
const path = require('node:path');
const { installMetadataAdvisoryRouting, prepareMetadata } = require('./build-metadata-preload.cjs');
installMetadataAdvisoryRouting();
const entry = process.argv[1] && path.basename(process.argv[1]);
if (entry === 'gulp' || entry === 'gulp.js') prepareMetadata();
