const fs = require('node:fs');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const root = path.resolve(__dirname, '..');
const destination = path.join(root, 'artifacts');
fs.mkdirSync(destination,{recursive:true});
execFileSync('npm',['pack','--pack-destination',destination],{cwd:path.join(root,'packages/spfx-react-toolkit'),stdio:'inherit'});
