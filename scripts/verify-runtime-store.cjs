const assert = require('assert');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const root = path.resolve(__dirname, '..');
const outDir = path.join(os.tmpdir(), 'spfx-runtime-store-verify');

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

execFileSync(
  'npx',
  [
    'tsc',
    'src/core/runtime-store.internal.ts',
    '--module',
    'commonjs',
    '--target',
    'es2020',
    '--skipLibCheck',
    '--esModuleInterop',
    '--outDir',
    outDir,
    '--pretty',
    'false'
  ],
  { cwd: root, stdio: 'inherit' }
);

const {
  createSPFxRuntimeStore,
  createDefaultSPFxRuntimeState
} = require(path.join(outDir, 'runtime-store.internal.js'));

const defaultState = createDefaultSPFxRuntimeState();
assert.deepStrictEqual(defaultState.teams, { supported: false, initialized: false });
assert.strictEqual(defaultState.theme, undefined);
assert.strictEqual(defaultState.displayMode, undefined);

const store = createSPFxRuntimeStore(defaultState);
assert.deepStrictEqual(store.getState(), defaultState);

let notifications = 0;
const unsubscribe = store.subscribe(() => {
  notifications += 1;
});

store.setState({ theme: undefined });
assert.strictEqual(notifications, 0, 'unchanged partial state must not notify');

const properties = { title: 'Hello' };
store.setState({ properties });
assert.strictEqual(store.getState().properties, properties);
assert.strictEqual(notifications, 1, 'changed partial state must notify once');

store.setState(previous => previous);
assert.strictEqual(notifications, 1, 'same updater state must not notify');

let snapshotNotifications = 0;
const unsubscribeDuringNotify = store.subscribe(() => {
  snapshotNotifications += 1;
  unsubscribeDuringNotify();
});

store.setState({ displayMode: 1 });
assert.strictEqual(snapshotNotifications, 1, 'listener should run during snapshot notification');
store.setState({ displayMode: 2 });
assert.strictEqual(snapshotNotifications, 1, 'unsubscribed listener should not run again');
assert.strictEqual(notifications, 3, 'remaining listener should receive both display mode updates');

unsubscribe();
store.setState({ containerSize: { width: 10, height: 20 } });
assert.strictEqual(notifications, 3, 'unsubscribed listener should not receive later updates');

console.log('runtime store verification passed');
