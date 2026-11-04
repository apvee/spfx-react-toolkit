const http = require('node:http');
const fs = require('node:fs/promises');
const path = require('node:path');

const contentTypes = {
  '.html': 'text/html; charset=utf-8', '.js': 'application/javascript',
  '.json': 'application/json', '.map': 'application/json', '.css': 'text/css',
  '.png': 'image/png', '.ico': 'image/x-icon', '.svg': 'image/svg+xml',
};
const owns = (object, key) => Object.prototype.hasOwnProperty.call(object, key);
const contained = (directory, file) => file.startsWith(directory + path.sep);

/** Serve a bounded fixture directory, or only the finite pathname-to-relative-file mapping.
 * Routes override files (including favicon's default 204). Failed reads return their
 * diagnostic with 404 (missing) or 500; invalid routes return 500 with the error.
 * Listen errors reject acquisition. close is awaited and idempotent.
 */
async function startFixtureServer({ directory, files, routes = {}, port = 0 }) {
  const root = await fs.realpath(directory);
  const server = http.createServer(async (request, response) => {
    try {
      const pathname = decodeURIComponent(new URL(request.url, 'http://127.0.0.1').pathname);
      if (owns(routes, pathname)) {
        const { status = 200, contentType, body = '' } = routes[pathname];
        if (typeof body !== 'string' && !Buffer.isBuffer(body)) throw new TypeError(`Invalid route body for ${pathname}`);
        response.writeHead(status, contentType ? { 'Content-Type': contentType } : {});
        response.end(body);
        return;
      }
      if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
      if (files && !owns(files, pathname)) { response.writeHead(404).end(); return; }
      const relative = files ? files[pathname] : pathname === '/' ? 'index.html' : pathname.slice(1);
      const file = path.resolve(root, relative);
      if (!contained(root, file)) { response.writeHead(403).end(); return; }
      // Resolve each path component before reading, including dangling symlink targets.
      let current = root;
      for (const part of path.relative(root, file).split(path.sep)) {
        current = path.join(current, part);
        const stat = await fs.lstat(current);
        if (stat.isSymbolicLink()) {
          current = path.resolve(path.dirname(current), await fs.readlink(current));
          if (!contained(root, current)) { response.writeHead(403).end(); return; }
          current = await fs.realpath(current);
          if (!contained(root, current)) { response.writeHead(403).end(); return; }
        }
      }
      const data = await fs.readFile(current);
      response.writeHead(200, { 'Content-Type': contentTypes[path.extname(file)] || 'application/octet-stream' });
      response.end(data);
    } catch (error) {
      response.writeHead(error.code === 'ENOENT' ? 404 : 500, { 'Content-Type': 'text/plain; charset=utf-8' });
      response.end(error.stack || String(error));
    }
  });
  await new Promise((resolve, reject) => {
    const onError = error => { server.removeListener('listening', onListening); reject(error); };
    const onListening = () => { server.removeListener('error', onError); resolve(); };
    server.once('error', onError);
    server.once('listening', onListening);
    server.listen(port, '127.0.0.1');
  });
  let closing;
  return {
    url: `http://127.0.0.1:${server.address().port}`,
    close() {
      if (!closing) closing = new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
      return closing;
    },
  };
}
module.exports = { startFixtureServer };
