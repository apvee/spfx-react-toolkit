const fs = require('node:fs');
const readdir = fs.readdirSync;
fs.readdirSync = function (...args) { return readdir.apply(this, args).reverse(); };
const ts = require('typescript');
const readDirectory = ts.sys.readDirectory;
ts.sys.readDirectory = (...args) => readDirectory(...args).reverse();
