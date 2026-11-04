const fs = require('node:fs');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const packageRoot = path.resolve(__dirname, '..');
const root = path.resolve(packageRoot, '../..');
const destination = path.join(root, 'artifacts');
fs.mkdirSync(destination,{recursive:true});
execFileSync('npm',['pack','--pack-destination',destination],{cwd:packageRoot,stdio:'inherit'});
