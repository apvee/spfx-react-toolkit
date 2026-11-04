const fs = require('node:fs');
const path = require('node:path');

/** @typedef {Record<string, string>} ArtifactSet Repository-relative paths to canonical UTF-8/LF text. */
/** @typedef {{missing: string[], changed: string[], orphan: string[]}} ArtifactDiff */
/** @typedef {{ok: boolean, mode: 'write'|'check', diff: ArtifactDiff}} ArtifactResult */
/** @typedef {{allows: (route: string) => boolean, owns: (route: string, text: string) => boolean, candidateRoots: string[], replaceUnownedExpected: boolean, allowTrailingBlankLines?: boolean}} ArtifactPolicy */
const compare = (a, b) => a < b ? -1 : a > b ? 1 : 0;
function validRoute(route) {
  return typeof route === 'string' && route.length > 0 && !/[\\\0?#:]/.test(route) &&
    !path.posix.isAbsolute(route) && route.split('/').every(part => part && part !== '.' && part !== '..');
}
/** Root aliases are permitted; symlinks below it are never output destinations. */
function checkedPath(root, route) {
  if (!validRoute(route)) throw new Error(`Invalid documentation path: ${route}.`);
  let current = root;
  const parts = route.split('/');
  for (const [index, part] of parts.entries()) {
    current = path.join(current, part);
    let stat;
    try { stat = fs.lstatSync(current); }
    catch (error) { if (error.code === 'ENOENT') continue; throw error; }
    if (stat.isSymbolicLink()) throw new Error(`Documentation path contains a symbolic link: ${route}.`);
    if (index < parts.length - 1 && !stat.isDirectory()) throw new Error(`Documentation ancestor is not a directory: ${route}.`);
  }
  return current;
}
function rootPath(root) {
  const absolute = fs.realpathSync(path.resolve(root));
  if (!fs.statSync(absolute).isDirectory()) throw new Error('Documentation root must be a directory.');
  return absolute;
}
/** Validate every output before comparisons, staging or mutations. */
function validateExpected(root, expected, policy) {
  if (!expected || typeof expected !== 'object' || Array.isArray(expected)) throw new Error('Expected documentation artifacts must be a path/text record.');
  const routes = Object.keys(expected).sort(compare);
  const portable = new Set();
  const inodes = new Set();
  for (const [route, content] of Object.entries(expected)) {
    if (!validRoute(route) || !policy.allows(route)) throw new Error(`Documentation destination is not allowed: ${route}.`);
    const folded = route.normalize('NFC').toLowerCase();
    if (portable.has(folded)) throw new Error(`Documentation destination collision: ${route}.`);
    for (const other of routes) if (other !== route && other.normalize('NFC').toLowerCase().startsWith(folded + '/')) throw new Error(`Documentation destination/ancestor collision: ${route} and ${other}.`);
    portable.add(folded);
    if (typeof content !== 'string' || content.includes('\r') || !content.endsWith('\n') || (!policy.allowTrailingBlankLines && content.endsWith('\n\n')) || Buffer.from(content, 'utf8').toString('utf8') !== content) throw new Error(`Documentation artifact requires canonical UTF-8/LF and one final newline: ${route}.`);
    const absolute = checkedPath(root, route);
    if (fs.existsSync(absolute)) {
      const stat = fs.lstatSync(absolute);
      if (!stat.isFile()) throw new Error(`Documentation destination is not a regular file: ${route}.`);
      const inode = `${stat.dev}:${stat.ino}`;
      if (inodes.has(inode)) throw new Error(`Documentation destinations share a file: ${route}.`);
      inodes.add(inode);
      if (!policy.replaceUnownedExpected && !policy.owns(route, fs.readFileSync(absolute, 'utf8'))) throw new Error(`Refusing to replace unowned manual artifact: ${route}.`);
    }
  }
}
/** Discover admitted candidates without following symbolic links. */
function candidates(root, policy) {
  const result = new Set();
  function walk(route, candidateRoot = false) {
    const absolute = checkedPath(root, route);
    if (!fs.existsSync(absolute)) return;
    const stat = fs.lstatSync(absolute);
    if (candidateRoot && policy.allows(route) && !stat.isFile()) throw new Error(`Generated artifact destination is not a regular file: ${route}.`);
    if (candidateRoot && !policy.allows(route) && !stat.isDirectory()) throw new Error(`Generated artifact candidate root is not a directory: ${route}.`);
    if (stat.isFile()) {
      if (policy.allows(route)) result.add(route);
      return;
    }
    if (!stat.isDirectory()) throw new Error(`Generated artifact destination is not a regular file: ${route}.`);
    for (const entry of fs.readdirSync(absolute, { withFileTypes: true }).sort((a, b) => compare(a.name, b.name))) {
      const child = route + '/' + entry.name;
      if (/(?:^|\/)(?:\.docs|_sources|node_modules|\.git)(?:\/|$)/.test(child)) continue;
      if (entry.isDirectory()) walk(child);
      else if (entry.isFile() && policy.allows(child)) result.add(child);
    }
  }
  for (const route of policy.candidateRoots) walk(route, true);
  return [...result].sort(compare);
}
/** Validate all destinations and ownership conflicts without writes or directory creation.
 * @param {string} root
 * @param {ArtifactSet} expected
 * @param {ArtifactPolicy} policy
 * @returns {void}
 */
function preflightGeneratedArtifacts(root, expected, policy) {
  root = rootPath(root);
  validateExpected(root, expected, policy);
  candidates(root, policy);
}
/**
 * Compare exact expected UTF-8 bytes and find only owned orphans in admitted roots.
 * Reads output bytes for comparison/ownership only; never writes or creates paths.
 * @param {string} root Repository directory.
 * @param {ArtifactSet} expected Complete expected output set.
 * @param {ArtifactPolicy} policy Destination and ownership rules.
 * @returns {ArtifactDiff} Ordinally sorted missing/changed/orphan paths.
 */
function compareGeneratedArtifacts(root, expected, policy) {
  root = rootPath(root);
  validateExpected(root, expected, policy);
  const diff = { missing: [], changed: [], orphan: [] };
  for (const route of Object.keys(expected).sort(compare)) {
    const absolute = checkedPath(root, route);
    if (!fs.existsSync(absolute)) diff.missing.push(route);
    else if (!fs.readFileSync(absolute).equals(Buffer.from(expected[route], 'utf8'))) diff.changed.push(route);
  }
  for (const route of candidates(root, policy)) if (!Object.hasOwn(expected, route) && policy.owns(route, fs.readFileSync(checkedPath(root, route), 'utf8'))) diff.orphan.push(route);
  return diff;
}
/**
 * Check is strictly read-only, including failed checks and directory mtimes.
 * Write stages all changed/missing file bytes first, then replaces each file and
 * removes owned orphans. Unchanged files keep their mtimes. Individual renames
 * are atomic on the repository filesystem; intermediate I/O failure can leave
 * earlier replacements applied, so this is not an all-files transaction. The
 * checkout must not be concurrently modified while applying the set.
 * @param {string} root Repository directory.
 * @param {ArtifactSet} expected Validated complete output set.
 * @param {'write'|'check'} mode Operation.
 * @param {ArtifactPolicy} policy Destination and ownership rules.
 * @returns {ArtifactResult} Freshness in check; applied original diff in write.
 */
function applyGeneratedArtifacts(root, expected, mode, policy) {
  if (mode !== 'write' && mode !== 'check') throw new Error(`Unknown documentation artifact mode: ${mode}.`);
  root = rootPath(root);
  const diff = compareGeneratedArtifacts(root, expected, policy);
  const stale = diff.missing.length + diff.changed.length + diff.orphan.length > 0;
  if (mode === 'check' || !stale) return { ok: !stale, mode, diff };
  const stage = fs.mkdtempSync(path.join(root, '.documentation-stage-'));
  try {
    const writes = [...diff.missing, ...diff.changed].sort(compare);
    for (const [index, route] of writes.entries()) fs.writeFileSync(path.join(stage, String(index)), expected[route], { encoding: 'utf8', flag: 'wx' });
    for (const [index, route] of writes.entries()) {
      const destination = checkedPath(root, route);
      fs.mkdirSync(path.dirname(destination), { recursive: true });
      // Recheck after directory creation and before each replacement/removal.
      fs.renameSync(path.join(stage, String(index)), checkedPath(root, route));
    }
    for (const route of diff.orphan) fs.unlinkSync(checkedPath(root, route));
  } finally { fs.rmSync(stage, { recursive: true, force: true }); }
  return { ok: true, mode, diff };
}
module.exports = { compareGeneratedArtifacts, preflightGeneratedArtifacts, applyGeneratedArtifacts, validRoute, checkedPath, rootPath };
