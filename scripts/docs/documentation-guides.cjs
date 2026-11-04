const fs = require('node:fs');
const path = require('node:path');
const MarkdownIt = require('markdown-it');
const defaultConfig = require('./documentation.config.cjs');
const { getDocumentationSelection, getDocumentationExample, resolveDocumentationReference } = require('./documentation-contract.cjs');
const { renderDocumentationExample } = require('./documentation-render.cjs');
const { markdownLabel, transformMarkdown, resolveMarkdown: resolveMarkdownProse, markdownAnchors, markdownLinks, relativeMarkdownTarget, editorialRoute, externalTarget } = require('./documentation-markdown.cjs');

const resolveMarkdown = (source, model, sourceFile, destination, include, publicOrigin) =>
  resolveMarkdownProse(source, sourceFile, destination, reference => resolveDocumentationReference(model, reference), include, publicOrigin);

/** @typedef {import('./documentation-model.cjs').Config} Config */
/** @typedef {import('./documentation-contract.cjs').DocumentationModel} DocumentationModel */
/** @typedef {import('./documentation-model.cjs').Diagnostic} Diagnostic */
/** @typedef {Record<string, string>} ArtifactSet */
/** @typedef {{source: string, route: string, content: string, title: string, summary: string, area: string, inventory?: {schemaVersion: 1, resources: {id: string, path: string, locales: string[]}[]}}} EditorialSource */
/** @typedef {{route: string, kind: 'markdown' | 'file', anchors: string[], content?: string}} PublicTarget */
const compare = (a, b) => a < b ? -1 : a > b ? 1 : 0;
const finish = text => text.replace(/\r\n?/g, '\n').replace(/\n*$/, '\n');
const privateRoute = route => /(?:^|\/)(?:\.docs|_sources)(?:\/|$)/.test(route);
/** Resolve existing ancestors before any local input read/target acceptance. */
function checkedInput(root, route, allowAuthoringSources = false) {
  const rootPath = path.resolve(root);
  const rootReal = fs.realpathSync(rootPath);
  const candidate = path.resolve(rootPath, route);
  const outside = relative => relative === '..' || relative.startsWith('..' + path.sep) || path.isAbsolute(relative);
  const fail = reason => {
    const message = `Documentation input ${route} is outside allowed public inputs: ${reason}.`;
    throw Object.assign(new Error(message), { code: 'DOC_OUTSIDE_INPUT', file: route, line: 1 });
  };
  if (outside(path.relative(rootPath, candidate))) fail('outside repository root');
  let ancestor = candidate;
  const missing = [];
  let ancestorReal;
  for (;;) {
    try { ancestorReal = fs.realpathSync(ancestor); break; }
    catch (error) {
      if (error.code !== 'ENOENT' && error.code !== 'ENOTDIR') throw error;
      missing.unshift(path.basename(ancestor));
      const parent = path.dirname(ancestor);
      if (parent === ancestor) throw error;
      ancestor = parent;
    }
  }
  const absolute = path.resolve(ancestorReal, ...missing);
  const relative = path.relative(rootReal, absolute);
  if (outside(relative)) fail('outside repository root through a symbolic link');
  const canonicalRoute = relative.split(path.sep).join('/');
  if (/(?:^|\/)\.docs(?:\/|$)/.test(canonicalRoute)) fail('private evidence');
  if (!allowAuthoringSources && /(?:^|\/)_sources(?:\/|$)/.test(canonicalRoute)) fail('private authoring source');
  return { absolute, route: canonicalRoute };
}
function files(root, route, publicOnly = false) {
  if (publicOnly && privateRoute(route)) return [];
  const directory = checkedInput(root, route, !publicOnly).absolute;
  if (!fs.existsSync(directory)) return [];
  return fs.readdirSync(directory, { withFileTypes: true }).sort((a, b) => compare(a.name, b.name)).flatMap(entry => {
    const file = route + '/' + entry.name;
    return entry.isDirectory() ? files(root, file, publicOnly) : entry.isFile() ? [file] : [];
  });
}
function localizationInventory(root, config) {
  const file = path.posix.join(config.packageRoot || defaultConfig.packageRoot, 'localization/resources.json');
  const inventory = JSON.parse(fs.readFileSync(checkedInput(root, file).absolute, 'utf8'));
  const ids = new Set();
  if (inventory.schemaVersion !== 1 || !Array.isArray(inventory.resources) || !inventory.resources.length) throw new Error(`Unsupported localization inventory schema: ${file}.`);
  for (const resource of inventory.resources) {
    if (!resource || typeof resource.id !== 'string' || !/^[A-Za-z_$][\w$]*$/.test(resource.id) || ids.has(resource.id) ||
      typeof resource.path !== 'string' || !/^lib\/(?:[a-z0-9-]+\/)+loc\/\{locale\}\.js$/.test(resource.path) ||
      !Array.isArray(resource.locales) || !resource.locales.length || resource.locales.some(locale => typeof locale !== 'string' || !/^[A-Za-z]{2,3}(?:-[A-Za-z0-9]+)+$/.test(locale)) || new Set(resource.locales).size !== resource.locales.length) {
      throw new Error(`Invalid localization inventory id/path/locales schema: ${file}.`);
    }
    ids.add(resource.id);
  }
  return inventory;
}
/**
 * Discover INDEX/getting-started/guides/development source Markdown by convention. H1 and
 * the first following prose paragraph supply public metadata. No output reads.
 * @param {string} root Repository directory.
 * @param {Config} config Project constants.
 * @returns {EditorialSource[]} Sorted source-owned pages.
 */
function discoverEditorialSources(root, config = defaultConfig) {
  const routes = files(root, 'docs/_sources').filter(file => /^docs\/_sources\/(?:INDEX\.md|(?:getting-started|guides|development)\/.+\.md)$/.test(file));
  return routes.map(source => {
    const content = fs.readFileSync(checkedInput(root, source, true).absolute, 'utf8').replace(/\r\n?/g, '\n');
    const tokens = new MarkdownIt({ html: true }).parse(content, {});
    const titleIndex = tokens.findIndex(token => token.type === 'heading_open' && token.tag === 'h1');
    if (titleIndex < 0 || !tokens[titleIndex + 1].content.trim()) throw new Error(`Missing H1 title in ${source}.`);
    let summary;
    for (let i = titleIndex + 3; i < tokens.length; i++) {
      if (tokens[i].type === 'heading_open') break;
      if (tokens[i].type === 'paragraph_open') { summary = tokens[i + 1].content.trim(); break; }
    }
    if (!summary) throw new Error(`Missing introduction in ${source}.`);
    // Snapshot the only supported finite inventory input while source discovery
    // still has an explicit root; rendering stays pure and root-independent.
    let needsInventory = false;
    transformMarkdown(content, { reference: reference => reference,
      include: directive => { if (/\binventory\s*=/.test(directive)) needsInventory = true; return directive; } });
    return { source, route: editorialRoute(source), title: tokens[titleIndex + 1].content.trim(), summary,
      area: source === 'docs/_sources/INDEX.md' ? 'Documentation' : source.split('/')[2], content,
      ...(needsInventory ? { inventory: localizationInventory(root, config) } : {}) };
  });
}
function attributes(directive) {
  const match = directive.match(/^<!--\s*toolkit:include\s+([\s\S]*?)\s*-->$/);
  if (!match) throw new Error('Invalid toolkit:include syntax.');
  const result = {};
  let text = match[1];
  while (text.trim()) {
    const attribute = text.match(/^\s*([a-z]+)="([^"\n]*)"(?=\s|$)/);
    if (!attribute || !['symbol', 'section', 'title', 'inventory'].includes(attribute[1]) || Object.hasOwn(result, attribute[1])) throw new Error('Unknown, duplicate or invalid include attribute syntax.');
    result[attribute[1]] = attribute[2];
    text = text.slice(attribute[0].length);
  }
  return result;
}
function includeFragment(directive, source, model) {
  const attrs = attributes(directive);
  if (attrs.inventory !== undefined) {
    if (Object.keys(attrs).length !== 1 || attrs.inventory !== 'localization.resources' || !source.inventory) throw new Error('Invalid or unavailable include inventory.');
    return '| Resource ID | Package-relative path |\n| --- | --- |\n' + [...source.inventory.resources].sort((a, b) => compare(a.id, b.id))
      .map(resource => `| \`${resource.id}\` | \`${resource.path}\` |`).join('\n');
  }
  if (!attrs.symbol || !['selection', 'example'].includes(attrs.section) || attrs.title !== undefined && attrs.section !== 'example') throw new Error('Include requires symbol and selection/example section; title is only valid for an example.');
  if (attrs.section === 'example') {
    const example = getDocumentationExample(model, attrs.symbol, attrs.title);
    return renderDocumentationExample({ ...example, title: example.title
      ? resolveMarkdown(example.title, model, example.source.file, source.route, undefined, resolveDocumentationReference(model, attrs.symbol).page?.route) : example.title }).replace(/\n$/, '');
  }
  const contract = getDocumentationSelection(model, attrs.symbol);
  const text = [contract.useWhen.length ? '### Use when\n\n' + contract.useWhen.map(item => '- ' + item).join('\n') : '',
    contract.avoidWhen.length ? '### Avoid when\n\n' + contract.avoidWhen.map(item => '- ' + item).join('\n') : ''].filter(Boolean).join('\n\n');
  // The outer transform patches original spans only. This independently
  // resolved fragment uses contract provenance and is never rebased twice.
  return resolveMarkdown(text, model, contract.source.file, source.route, undefined, resolveDocumentationReference(model, attrs.symbol).page?.route);
}
/** Render discovered editorial sources without reading files or writing output. */
function renderEditorialPages(sources, model, config = defaultConfig) {
  const artifacts = {};
  const destinations = new Set(model.pages.map(page => page.route));
  for (const source of [...sources].sort((a, b) => compare(a.route, b.route))) {
    if (destinations.has(source.route)) throw new Error(`Editorial destination collision: ${source.route}.`);
    destinations.add(source.route);
    let content = resolveMarkdown(source.content, model, source.source, source.route, directive => includeFragment(directive, source, model));
    if (source.route === 'docs/INDEX.md') {
      for (const [area, title] of [['getting-started', 'Getting started'], ['guides', 'Guides']]) {
        const pages = sources.filter(item => item.area === area).sort((a, b) => compare(a.route, b.route) || compare(a.title, b.title));
        if (pages.length) content += '\n\n## ' + title + '\n\n' + pages.map(page => `- [${markdownLabel(resolveMarkdown(page.title, model, page.source, source.route))}](${relativeMarkdownTarget(source.route, page.route)})${page.summary ? ': ' + resolveMarkdown(page.summary, model, page.source, source.route) : ''}`).join('\n');
      }
      content += '\n\n## Reference and development\n\n' + [
        ['API catalog', 'docs/api/INDEX.md'], ['Development index', 'docs/development/INDEX.md'],
        ['SharePoint validation index', 'docs/development/sharepoint-validation/INDEX.md'],
      ].map(([title, route]) => `- [${title}](${relativeMarkdownTarget(source.route, route)})`).join('\n');
    }
    artifacts[source.route] = finish(content);
  }
  return artifacts;
}
/**
 * Build an explicit virtual target universe. Existing generated-route Markdown
 * never becomes valid because it happens to remain on disk. Development pages
 * use the same virtual ownership boundary; local files are existence checked.
 * @param {string} root Repository directory.
 * @param {ArtifactSet} artifacts Expected virtual outputs.
 * @param {Config} config Project constants.
 * @returns {PublicTarget[]} Sorted verified targets.
 */
function collectPublicTargets(root, artifacts, config = defaultConfig) {
  const targets = new Map();
  const addMarkdown = (route, content) => targets.set(route, { route, kind: 'markdown', content, anchors: markdownAnchors(content) });
  for (const [route, content] of Object.entries(artifacts)) if (!privateRoute(route)) addMarkdown(route, content);
  // GitHub directory URLs render their README. Alias only the supplied virtual
  // README here, never a stale generated README read from disk.
  for (const [route, content] of Object.entries(artifacts)) if (!privateRoute(route) && route.endsWith('/README.md')) {
    addMarkdown(path.posix.dirname(route), content);
  }
  for (const route of files(root, 'docs', true)) {
    if (privateRoute(route) || targets.has(route)) continue;
    const input = checkedInput(root, route);
    if (route.endsWith('.md')) {
      if (route === 'docs/INDEX.md' || /^docs\/(?:api|getting-started|guides|development)\//.test(route)) continue;
      addMarkdown(route, fs.readFileSync(input.absolute, 'utf8'));
    } else targets.set(route, { route, kind: 'file', anchors: [] });
  }
  // Discover only actual referenced local files outside docs. Never crawl
  // node_modules or treat old generated Markdown as a fallback target.
  for (const target of [...targets.values()]) if (target.content) for (const href of markdownLinks(target.content)) {
    const resolved = localTarget(target.route, href);
    if (!resolved || resolved.route === 'llms.txt' || targets.has(resolved.route) || privateRoute(resolved.route) || /^docs\/(?:api|getting-started|guides|development)\//.test(resolved.route) && resolved.route.endsWith('.md') || resolved.route === 'docs/INDEX.md') continue;
    const input = checkedInput(root, resolved.route);
    const virtual = targets.get(input.route);
    if (virtual) { targets.set(resolved.route, { ...virtual, route: resolved.route }); continue; }
    if (input.route === 'llms.txt' || input.route === 'docs/INDEX.md' ||
      /^docs\/(?:api|getting-started|guides|development)\/.+\.md$/.test(input.route) ||
      input.route === path.posix.join(config.packageRoot || defaultConfig.packageRoot, 'README.md')) continue;
    const absolute = input.absolute;
    if (!fs.existsSync(absolute)) continue;
    const stat = fs.lstatSync(absolute);
    if (stat.isDirectory()) { targets.set(resolved.route, { route: resolved.route, kind: 'file', anchors: [] }); continue; }
    if (!stat.isFile() || resolved.route === path.posix.join(config.packageRoot || defaultConfig.packageRoot, 'README.md')) continue;
    if (resolved.route.endsWith('.md')) addMarkdown(resolved.route, fs.readFileSync(absolute, 'utf8'));
    else targets.set(resolved.route, { route: resolved.route, kind: 'file', anchors: [] });
  }
  return [...targets.values()].sort((a, b) => compare(a.route, b.route));
}
function localTarget(from, href) {
  if (!href || externalTarget(href)) return undefined;
  const hash = href.indexOf('#');
  const pathAndQuery = hash < 0 ? href : href.slice(0, hash);
  const query = pathAndQuery.indexOf('?');
  let file;
  let anchor;
  try {
    file = decodeURIComponent(query < 0 ? pathAndQuery : pathAndQuery.slice(0, query));
    anchor = hash < 0 ? '' : decodeURIComponent(href.slice(hash + 1));
  } catch { return { route: '', anchor: '' }; }
  const route = file ? path.posix.normalize(file.startsWith('/') ? file.slice(1) : path.posix.join(path.posix.dirname(from), file)) : from;
  return { route: route.startsWith('../') ? '' : route, anchor };
}
/** Validate links against supplied real/virtual targets; no filesystem/network. */
function validatePublicLinks(artifacts, targets, config = defaultConfig) {
  const index = new Map(targets.map(target => [target.route, target]));
  const diagnostics = [];
  for (const [file, content] of Object.entries(artifacts)) for (const href of markdownLinks(content)) {
    const local = localTarget(file, href);
    if (!local) continue;
    const target = index.get(local.route);
    if (!target) diagnostics.push({ code: 'DOC_LINK_TARGET', file, line: 1, message: `Missing local link target: ${href}.` });
    else if (local.anchor && target.kind === 'markdown' && !target.anchors.includes(local.anchor)) diagnostics.push({ code: 'DOC_LINK_ANCHOR', file, line: 1, message: `Missing Markdown anchor: ${href}.` });
  }
  return diagnostics.sort((a, b) => compare(a.file, b.file) || a.line - b.line || compare(a.code, b.code) || compare(a.message, b.message));
}
module.exports = { discoverEditorialSources, renderEditorialPages, collectPublicTargets, validatePublicLinks };
