const path = require('node:path');
const { startFixtureServer } = require('../helpers/fixture-server.cjs');
const root = path.resolve(__dirname, '../../../temp/sx-browser');
const port = Number(process.env.SX_BROWSER_PORT || 4317);
// This command owns the listener; the SX checker only consumes its external URL.
(async () => {
  const server = await startFixtureServer({ directory: root, port });
  const close = () => server.close().catch(error => { console.error(error); process.exitCode = 1; });
  process.once('SIGINT', close);
  process.once('SIGTERM', close);
  console.log(`useSx fixture: ${server.url}`);
})().catch(error => { console.error(error); process.exitCode = 1; });
