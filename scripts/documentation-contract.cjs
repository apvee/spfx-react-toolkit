const path = require('node:path');
const { TSDocConfiguration, TSDocParser, TSDocTagDefinition, TSDocTagSyntaxKind } = require('@microsoft/tsdoc');
const MarkdownIt = require('markdown-it');
const profile = require('../packages/spfx-react-toolkit/tsdoc.json');
const { resolvePublicReference } = require('./documentation-model.cjs');
const { markdownHeadingSlug } = require('./documentation-markdown.cjs');

/** @typedef {import('./documentation-model.cjs').ApiGraph} ApiGraph */
/** @typedef {import('./documentation-model.cjs').DocLocation} DocLocation */
/** @typedef {import('./documentation-model.cjs').Diagnostic} Diagnostic */
/** @typedef {{ownerKey: string, title?: string, language: 'ts'|'tsx', code: string, source: DocLocation}} DocExample */
/** @typedef {{summary: string, useWhen: string[], avoidWhen: string[], remarks: {title: string, body: string}[], params: {name: string, description: string}[], typeParams: {name: string, description: string}[], returns?: string, throws: string[], defaultValue?: string, deprecated?: string, docPage?: string, examples: DocExample[], see: string[], source: DocLocation, inheritedFrom?: string}} DocContract */
/** @typedef {{route: string, title: string, introduction: string, entityKeys: string[], anchors: Record<string, string>, bindingAnchors: Record<string, string>, typeHeadingSlugs: string[]}} ApiPage */
/** @typedef {{sourceRoot: string}} Config */
/** @typedef {{graph: ApiGraph, contracts: Record<string, DocContract>, bindingContracts: Record<string, DocContract>, pages: ApiPage[], examples: DocExample[], diagnostics: Diagnostic[]}} DocumentationModel */

const compare = (a, b) => a < b ? -1 : a > b ? 1 : 0;
const sortDiagnostics = items => [...new Map(items.map(item => [JSON.stringify(item), item])).values()]
  .sort((a, b) => compare(a.file, b.file) || a.line - b.line || compare(a.code, b.code) || compare(a.message, b.message) || compare(a.symbol || '', b.symbol || ''));
const markdown = new MarkdownIt({ html: false, linkify: false, typographer: false });
const parserConfig = new TSDocConfiguration();
for (const tag of profile.tagDefinitions) parserConfig.addTagDefinition(new TSDocTagDefinition({ ...tag, syntaxKind: TSDocTagSyntaxKind.BlockTag }));
const parser = new TSDocParser(parserConfig);
const memberContractKey = (ownerKey, name) => `${ownerKey}::member:${name}`;
const bindingContractKey = binding => `${binding.specifier}#${[binding.exportedName, ...binding.memberPath].join('.')}`;
function memberGroups(entity) {
  return [{ ownerKey: entity.key, members: entity.members, variants: entity.variants },
    ...entity.signatures.map(signature => ({ ownerKey: signature.key + '::returns', ...signature.returns }))];
}
function contractMembers(group) {
  const result = [];
  const sources = new Map();
  for (const member of group.members) {
    const key = memberContractKey(group.ownerKey, member.name);
    sources.set(key, member.source);
    result.push({ key, member });
  }
  for (const variant of group.variants) for (const member of variant.members) {
    const key = memberContractKey(group.ownerKey, member.name);
    const source = sources.get(key);
    const variantKey = source && (source.file !== member.source.file || source.line !== member.source.line) ? `${key}::variant:${variant.text}` : key;
    if (!sources.has(variantKey)) {
      sources.set(variantKey, member.source);
      result.push({ key: variantKey, member });
    }
  }
  return result;
}
const absent = text => /\b(?:omit(?:ted|ting)?|undefined|not (?:provided|specified|supplied))\b/i.test(text);

function excerpts(node) {
  if (node.kind === 'Excerpt') return [node];
  return node.getChildNodes().flatMap(excerpts);
}
function nodeRange(node) {
  const ranges = excerpts(node).filter(item => !item.content.isEmpty()).map(item => item.content.getContainingTextRange());
  return ranges.length ? { pos: Math.min(...ranges.map(item => item.pos)), end: Math.max(...ranges.map(item => item.end)) } : undefined;
}
const offsetLocation = (raw, source, offset) => ({ file: source.file, line: source.line + raw.slice(0, offset).split('\n').length - 1 });
function contentText(raw, node) {
  const range = nodeRange(node);
  return range ? raw.slice(range.pos, range.end).replace(/\r\n?/g, '\n').replace(/^[\t ]*\* ?/gm, '').trim() : '';
}
function addDiagnostic(diagnostics, source, symbol, code, message) {
  diagnostics.push({ code, ...source, symbol, message });
}
function items(text) {
  const tokens = markdown.parse(text, {});
  const result = [];
  let depth = 0;
  let start;
  for (const token of tokens) {
    if (token.type === 'list_item_open') {
      if (depth === 0) start = token.map?.[0];
      depth += 1;
    } else if (token.type === 'list_item_close') {
      depth -= 1;
      if (depth === 0 && start !== undefined) {
        const opening = tokens.find(item => item.type === 'list_item_open' && item.map?.[0] === start);
        const body = text.split('\n').slice(start, opening.map[1]).join('\n').replace(/^\s*(?:[-+*]|\d+[.)])\s+/, '').trim();
        result.push(body);
      }
    }
  }
  return result.length ? result : text ? [text] : [];
}
function remarks(text, source, ownerKey, diagnostics) {
  const tokens = markdown.parse(text, {});
  const headings = tokens.filter(token => token.type === 'heading_open' && token.tag === 'h3');
  if (!headings.length) return text ? [{ title: '', body: text }] : [];
  const lines = text.split('\n');
  const sections = [];
  const preface = lines.slice(0, headings[0].map[0]).join('\n').trim();
  if (preface) sections.push({ title: '', body: preface });
  for (let index = 0; index < headings.length; index += 1) {
    const heading = headings[index];
    const inline = tokens[tokens.indexOf(heading) + 1];
    const title = inline.content.trim();
    const body = lines.slice(heading.map[1], headings[index + 1]?.map[0] ?? lines.length).join('\n').trim();
    if (!body) addDiagnostic(diagnostics, source, ownerKey, 'DOC_EMPTY_REMARK', `Remarks section "${title}" is empty.`);
    if (sections.some(section => section.title.toLowerCase() === title.toLowerCase())) addDiagnostic(diagnostics, source, ownerKey, 'DOC_DUPLICATE_REMARK', `Duplicate remarks section "${title}".`);
    sections.push({ title, body });
  }
  return sections;
}

/** Adapt Markdown fences to the pinned TSDoc grammar without changing offsets. */
function parserText(raw) {
  const original = raw.split('\n');
  const body = original.map((line, index) => line.replace(index === 0 ? /^\s*\/\*\* ?/ : /^\s*\* ?/, '').replace(/\*\/\s*$/, ''));
  const adapted = [...original];
  for (const token of markdown.parse(body.join('\n'), {})) {
    if (token.type !== 'fence') continue;
    const [start, end] = token.map;
    const rewriteDelimiter = line => line.replace(/[`~]{3,}/, run => '```' + ' '.repeat(run.length - 3));
    adapted[start] = rewriteDelimiter(adapted[start]);
    const last = end - 1;
    const hasClose = last > start && /^\s*([`~])\1{2,}\s*$/.test(body[last]);
    if (hasClose) adapted[last] = rewriteDelimiter(adapted[last]);
    for (let index = start + 1; index < (hasClose ? last : end); index += 1) {
      // An inner shorter backtick run is code in Markdown, but would close a
      // three-backtick TSDoc fence. Its original spelling stays in the model.
      adapted[index] = adapted[index].replace(/^(\s*\*?\s*)`{3}/, '$1~``');
    }
  }
  return adapted.join('\n');
}

/** Parse exactly once with the shared tag profile; Markdown only tokenizes block content. */
function parseContract(raw, source, ownerKey, diagnostics) {
  const contract = { summary: '', useWhen: [], avoidWhen: [], remarks: [], params: [], typeParams: [], throws: [], examples: [], see: [], source };
  if (!raw) return contract;
  const context = parser.parseString(parserText(raw));
  for (const message of context.log.messages) addDiagnostic(diagnostics,
    offsetLocation(raw, source, message.textRange.pos), ownerKey, message.messageId, message.unformattedText);
  const doc = context.docComment;
  // Modifier tags are AST siblings, but a summary excerpt range can span the
  // surrounding soft breaks and therefore include their original spelling.
  // Blank only parser-recognized modifiers at their exact offsets; literal
  // occurrences inside prose/code remain untouched and source offsets survive.
  let summaryRaw = raw;
  for (const modifier of doc.modifierTagSet.nodes) {
    const range = nodeRange(modifier);
    if (range) summaryRaw = summaryRaw.slice(0, range.pos) + ' '.repeat(range.end - range.pos) + summaryRaw.slice(range.end);
  }
  contract.summary = contentText(summaryRaw, doc.summarySection);
  const blocks = [...doc.customBlocks, ...doc.seeBlocks, ...doc.params.blocks, ...doc.typeParams.blocks,
    doc.remarksBlock, doc.returnsBlock, doc.deprecatedBlock].filter(Boolean);
  blocks.sort((a, b) => (nodeRange(a)?.pos ?? 0) - (nodeRange(b)?.pos ?? 0));
  const seen = new Set();
  for (const block of blocks) {
    const tag = block.blockTag.tagName;
    const text = contentText(raw, block.content);
    const blockSource = offsetLocation(raw, source, nodeRange(block)?.pos ?? 0);
    if (!text) addDiagnostic(diagnostics, blockSource, ownerKey, 'DOC_EMPTY_TAG', `${tag} requires content.`);
    const singleton = ['@docPage', '@defaultValue', '@default', '@returns', '@deprecated'];
    const canonicalTag = tag === '@default' ? '@defaultValue' : tag;
    if (singleton.includes(tag) && seen.has(canonicalTag)) addDiagnostic(diagnostics, blockSource, ownerKey, 'DOC_DUPLICATE_TAG', `Duplicate ${canonicalTag}.`);
    seen.add(canonicalTag);
    switch (tag) {
      case '@useWhen': contract.useWhen.push(...items(text)); break;
      case '@avoidWhen': contract.avoidWhen.push(...items(text)); break;
      case '@docPage': contract.docPage = text; break;
      case '@remarks': contract.remarks.push(...remarks(text, blockSource, ownerKey, diagnostics)); break;
      case '@param': contract.params.push({ name: block.parameterName, description: text }); break;
      case '@typeParam': contract.typeParams.push({ name: block.parameterName, description: text }); break;
      case '@template': {
        const match = /^(\w+)\s*(?:-\s*)?([\s\S]*)$/.exec(text);
        if (match) contract.typeParams.push({ name: match[1], description: match[2] });
        else addDiagnostic(diagnostics, blockSource, ownerKey, 'DOC_INVALID_TYPE_PARAM', '@template requires a generic name and description.');
        break;
      }
      case '@returns': contract.returns = text; break;
      case '@throws': contract.throws.push(text); break;
      case '@default': case '@defaultValue': contract.defaultValue = text; break;
      case '@deprecated': contract.deprecated = text; break;
      case '@see': contract.see.push(text); break;
      case '@example': {
        const tokens = markdown.parse(text, {});
        const fences = tokens.filter(token => token.type === 'fence');
        if (fences.length !== 1 || !['ts', 'tsx'].includes(fences[0]?.info.trim()) || !fences[0]?.content.trim()) {
          addDiagnostic(diagnostics, blockSource, ownerKey, 'DOC_INVALID_EXAMPLE', 'Each @example requires exactly one nonempty ts or tsx fenced code block.');
          break;
        }
        const fence = fences[0];
        const title = text.split('\n').slice(0, fence.map[0]).join('\n').trim();
        const range = nodeRange(block.content);
        const contentSource = offsetLocation(raw, source, range?.pos ?? 0);
        // Markdown offsets start after contentText trims leading blank margins.
        // Count those physical lines without trimming or rewriting fenced code.
        const untrimmed = range ? raw.slice(range.pos, range.end).replace(/\r\n?/g, '\n').replace(/^[\t ]*\* ?/gm, '') : '';
        const leading = untrimmed.slice(0, untrimmed.length - untrimmed.trimStart().length);
        contentSource.line += leading.split('\n').length - 1;
        contract.examples.push({ ownerKey, ...(title ? { title } : {}), language: fence.info.trim(), code: fence.content,
          source: { file: source.file, line: contentSource.line + fence.map[0] + 1 } });
        break;
      }
    }
  }
  if (doc.inheritDocTag) {
    // Keep the actual declared upstream reference. No implicit provider/default
    // claims are copied from the upstream prose into a local toolkit contract.
    contract.inheritedFrom = doc.inheritDocTag.declarationReference?.emitAsTsdoc().trim();
  }
  return contract;
}
function validPage(value, sourceFallback = false) {
  if (!sourceFallback) return /^[A-Za-z0-9_-]+(?:\/[A-Za-z0-9_-]+)*$/.test(value);
  return /^[A-Za-z0-9_-][A-Za-z0-9._-]*(?:\/[A-Za-z0-9_-][A-Za-z0-9._-]*)*$/.test(value) && !/\.(?:md|tsx?)$/i.test(value) && !value.split('/').some(part => part === '.' || part === '..');
}

/** Natural runtime anchors stay stable; colliding type-only headings state their kind visibly. */
function apiHeadingName(page, entity, name) {
  return ['interface', 'type'].includes(entity.kind) && page.typeHeadingSlugs?.includes(markdownHeadingSlug(name)) ? `${name} (type)` : name;
}

/**
 * Normalize all source-owned documentation and routes without rendering or IO.
 * The graph remains canonical; this layer neither extracts types nor inventories APIs.
 * @param {ApiGraph} graph Semantic source graph.
 * @param {Config} config Versioned source root used for fallback routes.
 * @returns {DocumentationModel} Serializable source contracts and stable diagnostics.
 */
function buildDocumentationModel(graph, config) {
  const contracts = {};
  const bindingContracts = {};
  const diagnostics = [...graph.diagnostics];
  const entities = new Map(graph.entities.map(entity => [entity.key, entity]));
  // Decide physical module ownership from the complete canonical graph before
  // any member can create its shared file contract through iteration order.
  const moduleOwners = new Map();
  for (const entity of graph.entities) {
    if (entity.kind !== 'namespace' || entity.namespaceRole !== 'sourceFile' || entity.visibility !== 'public' ||
      !entity.moduleDoc || entity.moduleDocSource?.file !== entity.source.file ||
      !graph.bindings.some(binding => binding.entityKey === entity.key && binding.specifier && binding.runtimeExport === true)) continue;
    const candidates = moduleOwners.get(entity.source.file) || new Map();
    candidates.set(entity.key, entity);
    moduleOwners.set(entity.source.file, candidates);
  }
  for (const [file, candidates] of moduleOwners) if (candidates.size > 1) {
    const source = [...candidates.values()].sort((a, b) => compare(a.key, b.key))[0].moduleDocSource;
    addDiagnostic(diagnostics, source, file, 'DOC_AMBIGUOUS_MODULE_OWNER', 'A module header requires one canonical public runtime namespace owner.');
  }
  for (const entity of graph.entities) {
    if (entity.visibility === 'external') continue;
    if (entity.moduleDoc && !contracts[entity.source.file]) {
      const candidates = moduleOwners.get(entity.source.file);
      const owner = candidates?.size === 1 ? [...candidates.values()][0] : undefined;
      const sameHeader = owner && owner.moduleDoc === entity.moduleDoc && owner.moduleDocSource.file === entity.moduleDocSource?.file &&
        owner.moduleDocSource.line === entity.moduleDocSource?.line;
      contracts[entity.source.file] = parseContract(entity.moduleDoc, entity.moduleDocSource || entity.source, sameHeader ? owner.key : entity.source.file, diagnostics);
    }
    contracts[entity.key] = entity.namespaceRole === 'sourceFile' && !entity.rawDoc && contracts[entity.source.file]
      ? contracts[entity.source.file]
      : parseContract(entity.rawDoc, entity.docSource || entity.source, entity.key, diagnostics);
    for (const signature of entity.signatures) {
      if (!signature.sourceOwned) {
        // Inferred external aliases retain the toolkit's local contract.
        contracts[signature.key] = contracts[entity.key];
      } else if (signature.rawDoc) {
        contracts[signature.key] = parseContract(signature.rawDoc, signature.docSource || signature.source, signature.key, diagnostics);
      } else if (entity.signatures.length === 1) {
        contracts[signature.key] = contracts[entity.key];
      } else {
        // A shared purpose may describe overloads, but a different overload's
        // params/generics/returns are never an implicit signature contract.
        contracts[signature.key] = { ...parseContract('', signature.source, signature.key, diagnostics), summary: contracts[entity.key].summary };
      }
    }
    for (const group of memberGroups(entity)) {
      for (const { member, key } of contractMembers(group)) {
        if (!contracts[key]) contracts[key] = parseContract(member.rawDoc, member.docSource || member.source, key, diagnostics);
      }
    }
  }
  for (const entity of graph.entities) {
    if (!entity.parameterDocumentation || entity.rawDoc) continue;
    const { ownerKey, name } = entity.parameterDocumentation;
    const owner = contracts[ownerKey];
    const parameter = owner?.params.find(item => item.name === name);
    if (parameter) contracts[entity.key] = { ...contracts[entity.key], summary: parameter.description, source: owner.source };
  }
  for (const binding of graph.bindings) {
    if (binding.rawDoc && contracts[binding.entityKey] &&
      (binding.rawDoc !== entities.get(binding.entityKey).rawDoc || binding.docSource?.file !== entities.get(binding.entityKey).docSource?.file)) {
      const key = bindingContractKey(binding);
      bindingContracts[key] = parseContract(binding.rawDoc, binding.docSource || entities.get(binding.entityKey).source, key, diagnostics);
    }
  }
  // Inherited member provenance comes from actual declarations and heritage,
  // never from name suffixes or a resource manifest.
  for (const entity of graph.entities) {
    for (const base of entity.heritage) for (const baseKey of base.references) {
      const upstream = entities.get(baseKey);
      if (!upstream || upstream.visibility === 'external') continue;
      for (const member of entity.members) {
        const original = upstream.members.find(item => item.name === member.name);
        if (!original) continue;
        const contract = contracts[memberContractKey(entity.key, member.name)];
        if (contract) {
          contract.inheritedFrom = memberContractKey(upstream.key, member.name);
          if (original.source.file === member.source.file && original.source.line === member.source.line) {
            delete contract.defaultValue;
            contract.remarks = [];
          }
        }
      }
    }
  }
  for (const entity of graph.entities) for (const group of memberGroups(entity)) {
    for (const { member, key } of contractMembers(group)) {
      if (!member.ownerKey || member.ownerKey === entity.key) continue;
      const upstreamKey = memberContractKey(member.ownerKey, member.name);
      const contract = contracts[key];
      if (!contract || !contracts[upstreamKey]) continue;
      contract.inheritedFrom = upstreamKey;
      delete contract.defaultValue;
      contract.remarks = [];
    }
  }
  const directKeys = new Set(graph.bindings.filter(binding => !binding.memberPath.length).map(binding => binding.entityKey));
  const selectionDone = new Set();
  function inheritSelection(key, visiting = new Set()) {
    if (selectionDone.has(key) || visiting.has(key)) return;
    visiting.add(key);
    const contract = contracts[key];
    if (!contract || directKeys.has(key)) return;
    const owners = [...new Set(graph.bindings.filter(binding => binding.entityKey === key && binding.ownerKey).map(binding => binding.ownerKey))];
    if (owners.length === 1) {
      inheritSelection(owners[0], visiting);
      const owner = contracts[owners[0]];
      if (owner) {
        if (!contract.useWhen.length) contract.useWhen = owner.useWhen;
        if (!contract.avoidWhen.length) contract.avoidWhen = owner.avoidWhen;
      }
    }
    selectionDone.add(key);
  }
  for (const entity of graph.entities) inheritSelection(entity.key);
  for (const binding of graph.bindings) {
    if (!binding.ownerKey) continue;
    const key = bindingContractKey(binding);
    const local = bindingContracts[key];
    const own = local || contracts[binding.entityKey];
    const owner = contracts[binding.ownerKey];
    if (!own || !owner) continue;
    // If multiple public owners provide different contexts, preserve the context
    // on the particular binding instead of overwriting the shared declaration.
    if (local || (!own.useWhen.length && owner.useWhen.length) || (!own.avoidWhen.length && owner.avoidWhen.length)) {
      bindingContracts[key] = { ...own,
        useWhen: own.useWhen.length ? own.useWhen : owner.useWhen,
        avoidWhen: own.avoidWhen.length ? own.avoidWhen : owner.avoidWhen };
    }
  }
  const routes = new Map();
  function routeFor(entity, visiting = new Set()) {
    if (routes.has(entity.key)) return routes.get(entity.key);
    if (visiting.has(entity.key)) return undefined;
    visiting.add(entity.key);
    const contract = contracts[entity.key];
    if (!contract) return undefined;
    const own = contract.docPage || contracts[entity.source.file]?.docPage;
    const ownerKeys = [...new Set(graph.bindings.filter(binding => binding.entityKey === entity.key && binding.ownerKey).map(binding => binding.ownerKey))];
    const ownerRoutes = ownerKeys.map(key => entities.get(key)).filter(Boolean).map(owner => routeFor(owner, new Set(visiting))).filter(Boolean);
    const fallback = path.posix.relative(config.sourceRoot, entity.source.file).replace(/\.(?:tsx?|d\.ts)$/, '');
    const candidate = own || ownerRoutes[0] || fallback;
    if (!validPage(candidate, !own)) {
      addDiagnostic(diagnostics, contract.source, entity.key, 'DOC_INVALID_PAGE', `Invalid API page "${candidate}"; use a relative extension-free path without traversal.`);
      return undefined;
    }
    if (!own && new Set(ownerRoutes).size > 1) addDiagnostic(diagnostics, contract.source, entity.key, 'DOC_PAGE_CONFLICT', 'Public owners assign this entity to different pages.');
    routes.set(entity.key, candidate);
    return candidate;
  }
  for (const [key, contract] of Object.entries({ ...contracts, ...bindingContracts })) {
    if (contract.docPage !== undefined && !validPage(contract.docPage)) addDiagnostic(diagnostics, contract.source, key, 'DOC_INVALID_PAGE', `Invalid @docPage "${contract.docPage}".`);
  }
  for (const entity of graph.entities) if (entity.visibility !== 'external') routeFor(entity);
  for (const binding of graph.bindings) {
    const contract = bindingContracts[bindingContractKey(binding)];
    if (contract?.docPage && validPage(contract.docPage) && routes.get(binding.entityKey) !== contract.docPage) {
      addDiagnostic(diagnostics, contract.source, bindingContractKey(binding), 'DOC_PAGE_CONFLICT', `Binding @docPage "${contract.docPage}" conflicts with entity page "${routes.get(binding.entityKey)}".`);
    }
  }
  const pageMap = new Map();
  const portableRoutes = new Map();
  const preferredNames = new Map();
  for (const entity of graph.entities) {
    const route = routes.get(entity.key);
    if (!route) continue;
    const fullRoute = `docs/api/${route}.md`;
    const prior = portableRoutes.get(fullRoute.toLowerCase());
    if (prior && prior !== fullRoute) addDiagnostic(diagnostics, contracts[entity.key].source, entity.key, 'DOC_PAGE_COLLISION', `API routes collide across case-sensitive filesystems: ${prior} and ${fullRoute}.`);
    portableRoutes.set(fullRoute.toLowerCase(), fullRoute);
    if (!pageMap.has(fullRoute)) pageMap.set(fullRoute, { route: fullRoute, title: route.split('/').at(-1), introduction: '', entityKeys: [], anchors: {}, bindingAnchors: {}, typeHeadingSlugs: [] });
    const page = pageMap.get(fullRoute);
    const publicNames = graph.bindings.filter(binding => binding.entityKey === entity.key)
      .map(binding => ({ access: [binding.exportedName, ...binding.memberPath].join('.'), nested: binding.memberPath.length > 0 }))
      .sort((a, b) => Number(b.nested) - Number(a.nested) || a.access.length - b.access.length || compare(a.access, b.access));
    page.entityKeys.push(entity.key);
    // Keep the preferred spelling until all runtime/type names on the page are known.
    preferredNames.set(entity.key, publicNames[0]?.access || entity.qualifiedName);
  }
  for (const page of pageMap.values()) {
    page.entityKeys.sort(compare);
    page.introduction = getPageIntroductions({ graph, contracts }, page)[0]?.contract.summary || '';
    const names = page.entityKeys.flatMap(key => {
      const entity = entities.get(key);
      return [entity.qualifiedName, preferredNames.get(key), ...graph.bindings.filter(binding => binding.entityKey === key).map(binding => [binding.exportedName, ...binding.memberPath].join('.'))]
        .map(name => ({ name, typeOnly: ['interface', 'type'].includes(entity.kind) }));
    });
    const runtimeSlugs = new Set(names.filter(item => !item.typeOnly).map(item => markdownHeadingSlug(item.name)));
    page.typeHeadingSlugs = [...new Set(names.filter(item => item.typeOnly && runtimeSlugs.has(markdownHeadingSlug(item.name))).map(item => markdownHeadingSlug(item.name)))].sort(compare);
    const owners = new Map();
    for (const key of page.entityKeys) {
      const heading = apiHeadingName(page, entities.get(key), preferredNames.get(key));
      const anchor = markdownHeadingSlug(heading);
      if (owners.has(anchor) && owners.get(anchor) !== key) addDiagnostic(diagnostics, contracts[key].source, key, 'DOC_ANCHOR_COLLISION', `Duplicate anchor "${anchor}" on ${page.route}.`);
      owners.set(anchor, key);
      page.anchors[key] = anchor;
    }
    for (const binding of graph.bindings.filter(binding => page.entityKeys.includes(binding.entityKey))) {
      const key = bindingContractKey(binding);
      const anchor = markdownHeadingSlug(apiHeadingName(page, entities.get(binding.entityKey), [binding.exportedName, ...binding.memberPath].join('.')));
      const prior = owners.get(anchor);
      if (prior && prior !== binding.entityKey) {
        const contract = bindingContracts[key] || contracts[binding.entityKey];
        addDiagnostic(diagnostics, contract.source, key, 'DOC_ANCHOR_COLLISION', `Public binding anchor "${anchor}" collides on ${page.route}.`);
      }
      owners.set(anchor, binding.entityKey);
      page.bindingAnchors[key] = anchor;
    }
  }
  const pages = [...pageMap.values()].sort((a, b) => compare(a.route, b.route));
  const sourceExamples = new Map();
  for (const contract of Object.values({ ...contracts, ...bindingContracts })) {
    contract.examples = contract.examples.map(example => {
      const identity = JSON.stringify([example.source, example.title, example.language, example.code]);
      if (!sourceExamples.has(identity)) sourceExamples.set(identity, example);
      return sourceExamples.get(identity);
    });
  }
  const examples = [...sourceExamples.values()];
  return { graph, contracts, bindingContracts, pages, examples, diagnostics: sortDiagnostics(diagnostics) };
}

/**
 * Select the page's visible introductions once, with their original contracts.
 * Owning module introductions take precedence over child summaries; source-file
 * namespace contracts consumed by their own entity remain in that entity body.
 * Stable source/key order makes page rendering and navigation metadata agree.
 * @param {Pick<DocumentationModel, 'graph'|'contracts'>} model Canonical graph/contracts.
 * @param {ApiPage} page Destination page.
 * @returns {{key: string, contract: DocContract, full: boolean}[]} Visible module contracts, or the fallback summary contract.
 */
function getPageIntroductions(model, page) {
  const entities = page.entityKeys.map(key => model.graph.entities.find(entity => entity.key === key));
  const modules = [...new Set(entities.map(entity => entity.source.file))]
    .filter(file => model.contracts[file]?.summary).sort(compare);
  if (modules.length) return modules.filter(file => !entities.some(entity =>
    entity.source.file === file && entity.namespaceRole === 'sourceFile' && !entity.rawDoc))
    .map(file => ({ key: file, contract: model.contracts[file], full: true }));
  const key = [...page.entityKeys].sort(compare).find(key => model.contracts[key]?.summary);
  return key ? [{ key, contract: model.contracts[key], full: false }] : [];
}

/**
 * Return validation diagnostics; checks structure, never the truth of prose.
 * @param {DocumentationModel} model Normalized contracts.
 * @returns {Diagnostic[]} Relative, stable diagnostics sorted by file/line/code.
 */
function validateDocumentation(model) {
  const diagnostics = [...model.diagnostics];
  const report = (contract, key, code, message) => addDiagnostic(diagnostics, contract.source, key, code, message);
  const validateNames = (contract, key, field, actual, label, required = actual) => {
    const documented = contract[field];
    const seen = new Set();
    for (const item of documented) {
      if (!actual.some(parameter => parameter.name === item.name)) report(contract, key, `DOC_UNKNOWN_${label}`, `No declared ${label.toLowerCase()} named "${item.name}".`);
      if (seen.has(item.name)) report(contract, key, `DOC_DUPLICATE_${label}`, `Duplicate documentation for "${item.name}".`);
      if (!item.description) report(contract, key, 'DOC_EMPTY_TAG', `Description for "${item.name}" is empty.`);
      seen.add(item.name);
    }
    for (const parameter of required) {
      const item = documented.find(item => item.name === parameter.name);
      if (!item) report(contract, key, `DOC_MISSING_${label}`, `Document "${parameter.name}".`);
      else if (parameter.optional && !absent(item.description)) report(contract, key, 'DOC_MISSING_OMISSION', `Describe omission or undefined for optional input "${parameter.name}".`);
    }
  };
  for (const [key, contract] of Object.entries({ ...model.contracts, ...model.bindingContracts })) {
    if (contract.examples.length > 1) {
      const titles = new Set();
      for (const example of contract.examples) {
        if (!example.title) report(contract, key, 'DOC_MISSING_EXAMPLE_TITLE', 'Multiple examples require unique local titles.');
        else if (titles.has(example.title)) report(contract, key, 'DOC_DUPLICATE_EXAMPLE_TITLE', `Duplicate local example title "${example.title}".`);
        titles.add(example.title);
      }
    }
  }
  for (const entity of model.graph.entities) {
    if (entity.visibility === 'external') continue;
    const key = entity.key;
    const contract = model.contracts[key];
    if (!contract.summary) report(contract, key, 'DOC_MISSING_SUMMARY', 'Describe the API purpose and behavior.');
    const direct = model.graph.bindings.some(binding => binding.entityKey === key && binding.memberPath.length === 0);
    if (direct && !contract.examples.length) report(contract, key, 'DOC_MISSING_EXAMPLE', 'Public top-level APIs and types require a complete fenced example.');
    if (direct && !['type', 'interface'].includes(entity.kind) && !contract.useWhen.length) report(contract, key, 'DOC_MISSING_USE_WHEN', 'Directly selected runtime APIs require @useWhen.');
    const generics = [...new Map([...entity.typeParameters, ...entity.signatures.flatMap(signature => signature.typeParameters)].map(parameter => [parameter.name, parameter])).values()];
    validateNames(contract, key, 'typeParams', generics, 'TYPE_PARAM', entity.typeParameters);
    // A declaration comment shared with a source-owned signature describes that
    // signature, not the union of other overloads' names and optionality. Match
    // both comment text and provenance: distinct comments can share one line.
    // Separate entity overviews retain the aggregate callable-name policy.
    const allParameters = [...new Map(entity.signatures.flatMap(signature => signature.parameters).map(parameter => [parameter.name, parameter])).values()];
    const documentedSignature = entity.signatures.find(signature => signature.sourceOwned && signature.rawDoc === entity.rawDoc &&
      model.contracts[signature.key].source.file === contract.source.file && model.contracts[signature.key].source.line === contract.source.line);
    if (contract.params.length) validateNames(contract, key, 'params', documentedSignature ? documentedSignature.parameters : allParameters, 'PARAM');
    for (const signature of entity.signatures) {
      const own = model.contracts[signature.key];
      const effective = own;
      validateNames(effective, signature.key, 'params', signature.parameters, 'PARAM');
      // The retained comment belongs to the original local declaration, even
      // when this entity is a const with a closed instantiation of its interface.
      const lexical = signature.sourceOwned ? signature.lexicalTypeParameters : [];
      validateNames(effective, signature.key, 'typeParams', [...entity.typeParameters, ...lexical, ...signature.typeParameters], 'TYPE_PARAM', signature.typeParameters);
      if (!['void', 'undefined', 'never'].includes(signature.returns.text) && !effective.returns) report(effective, signature.key, 'DOC_MISSING_RETURNS', 'Describe the returned value.');
    }
    for (const group of memberGroups(entity)) for (const { member, key: memberKey } of contractMembers(group)) {
      const own = model.contracts[memberKey];
      if (!own) continue;
      if (!own.summary) report(own, memberKey, 'DOC_MISSING_SUMMARY', `Describe member "${member.name}".`);
      if (member.optional && !absent([own.summary, own.defaultValue, ...own.remarks.map(item => item.body)].filter(Boolean).join('\n'))) report(own, memberKey, 'DOC_MISSING_OMISSION', `Describe omission or undefined for optional member "${member.name}".`);
    }
  }
  for (const binding of model.graph.bindings) {
    const key = bindingContractKey(binding);
    const contract = model.bindingContracts[key];
    if (!contract) continue;
    const entity = model.graph.entities.find(item => item.key === binding.entityKey);
    if (!contract.summary) report(contract, key, 'DOC_MISSING_SUMMARY', 'Describe the binding purpose and behavior.');
    if (binding.memberPath.length === 0) {
      if (!contract.examples.length) report(contract, key, 'DOC_MISSING_EXAMPLE', 'Public binding contracts require a complete fenced example.');
      if (!['type', 'interface'].includes(entity.kind) && !contract.useWhen.length) report(contract, key, 'DOC_MISSING_USE_WHEN', 'Directly selected binding contracts require @useWhen.');
    }
    if (contract.params.length) validateNames(contract, key, 'params', entity.signatures.flatMap(signature => signature.parameters), 'PARAM');
    if (contract.typeParams.length) validateNames(contract, key, 'typeParams', [...entity.typeParameters, ...entity.signatures.flatMap(signature => signature.typeParameters)], 'TYPE_PARAM');
  }
  return sortDiagnostics(diagnostics);
}

/** Resolve a legal public access to its binding-local contract and shared route. */
function resolveDocumentationReference(model, reference) {
  const entity = resolvePublicReference(model.graph, reference);
  const separator = reference.indexOf('#');
  const specifier = separator === -1 ? undefined : reference.slice(0, separator);
  const access = separator === -1 ? reference : reference.slice(separator + 1);
  const binding = model.graph.bindings.find(item => item.entityKey === entity.key && (!specifier || item.specifier === specifier) && [item.exportedName, ...item.memberPath].join('.') === access);
  const contract = model.bindingContracts[bindingContractKey(binding)] || model.contracts[entity.key];
  const page = model.pages.find(item => item.entityKeys.includes(entity.key));
  return { entity, binding, contract, page, anchor: page?.bindingAnchors[bindingContractKey(binding)] || page?.anchors[entity.key] };
}
/** Both API rendering and guide inclusions consume these exact normalized fields. */
function getDocumentationSelection(model, reference) {
  return resolveDocumentationReference(model, reference).contract;
}
/** Select the unique example or an exact local title; never silently take the first. */
function getDocumentationExample(model, reference, title) {
  const { contract } = resolveDocumentationReference(model, reference);
  if (title !== undefined) {
    const matches = contract.examples.filter(example => example.title === title);
    if (matches.length === 1) return matches[0];
    throw new Error(`Missing or ambiguous example "${title}" for ${reference}.`);
  }
  if (contract.examples.length !== 1) throw new Error(`Select a unique example title for ${reference}.`);
  return contract.examples[0];
}

module.exports = { apiHeadingName, buildDocumentationModel, validateDocumentation, memberContractKey, bindingContractKey,
  resolveDocumentationReference, getDocumentationSelection, getDocumentationExample, getPageIntroductions };
