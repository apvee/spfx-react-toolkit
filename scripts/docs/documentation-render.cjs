const path = require('node:path');
const defaultConfig = require('./documentation.config.cjs');
const { getPageIntroductions, apiHeadingName, bindingContractKey, memberContractKey, resolveDocumentationReference } = require('./documentation-contract.cjs');
const { markdownHeadingSlug, markdownHeadingAnchor, transformMarkdown, resolveMarkdown: resolveMarkdownProse } = require('./documentation-markdown.cjs');

const resolveMarkdown = (source, model, sourceFile, destination, include, publicOrigin) =>
  resolveMarkdownProse(source, sourceFile, destination, reference => resolveDocumentationReference(model, reference), include, publicOrigin);

/** @typedef {import('./documentation-contract.cjs').DocumentationModel} DocumentationModel */
/** @typedef {import('./documentation-contract.cjs').DocContract} DocContract */
/** @typedef {import('./documentation-contract.cjs').DocExample} DocExample */
/** @typedef {import('./documentation-model.cjs').Config} Config */
/** @typedef {{route: string, title: string, summary: string, area: string}} PublicPage */
/** @typedef {Record<string, string>} ArtifactSet */

const compare = (left, right) => left < right ? -1 : left > right ? 1 : 0;
const accessName = binding => [binding.exportedName, ...binding.memberPath].join('.');
// Public names are literal text, so Markdown formatting cannot erase legal
// identifier characters (for example underscores) from canonical anchors.
const publicHeading = name => name.replace(/[\\`*_[\]<>]/g, '\\$&');
const finish = text => text.replace(/\r\n?/g, '\n').replace(/\n*$/, '\n');
const cell = text => String(text || '').replace(/\r\n?/g, '\n').replace(/\|/g, '\\|').replace(/\n/g, '<br>');
const codeSpan = text => {
  const delimiter = '`'.repeat(Math.max(1, ...[...text.matchAll(/`+/g)].map(match => match[0].length + 1)));
  return delimiter + (text.startsWith('`') || text.endsWith('`') ? ' ' + text + ' ' : text) + delimiter;
};
function codeBlock(code, language = 'ts') {
  const fence = '`'.repeat(Math.max(3, ...[...code.matchAll(/`+/g)].map(match => match[0].length + 1)));
  return fence + language + '\n' + code + (code.endsWith('\n') ? '' : '\n') + fence;
}
const editorialIdentity = contract => contract && JSON.stringify([
  contract.deprecated, contract.useWhen, contract.avoidWhen, contract.remarks,
  contract.defaultValue, contract.throws, contract.see,
  contract.examples.map(example => [example.title, example.language, example.code]),
]);
// Provenance is checked separately. Every visible normalized field participates;
// equal summaries or equal examples alone never establish contract identity.
const exampleIdentity = contract => JSON.stringify(contract.examples.map(example => [example.title, example.language, example.code]));
const contractIdentity = contract => JSON.stringify([
  contract.summary, contract.useWhen, contract.avoidWhen, contract.remarks,
  contract.params, contract.typeParams, contract.returns, contract.throws,
  contract.defaultValue, contract.deprecated, contract.see, exampleIdentity(contract),
]);
const sameSource = (left, right) => left.file === right.file && left.line === right.line;
const sameOwnedExamples = (left, right) => exampleIdentity(left) === exampleIdentity(right) && left.examples.every((example, position) =>
  example.ownerKey === right.examples[position].ownerKey && sameSource(example.source, right.examples[position].source));
const hasMemberDetail = contract => contract && (contract.remarks.length || contract.examples.length || contract.useWhen.length || contract.avoidWhen.length || contract.returns || contract.params.length || contract.typeParams.length || contract.throws.length || contract.see.length || contract.deprecated);
function repository(config = defaultConfig) {
  const url = (config.repositoryUrl || defaultConfig.repositoryUrl).replace(/\/$/, '');
  const ref = config.repositoryRef || defaultConfig.repositoryRef;
  return { blob: `${url}/blob/${ref}/`, raw: `${url.replace(/^https:\/\/github\.com\//, 'https://raw.githubusercontent.com/')}/${ref}/` };
}
function relativeTarget(from, route, anchor) {
  const relative = path.posix.relative(path.posix.dirname(from), route);
  return (relative || path.posix.basename(route)) + (anchor ? '#' + anchor : '');
}
function legalImport(binding, entity) {
  const type = ['interface', 'type'].includes(entity.kind) && !binding.memberPath.length ? ' type' : '';
  const statement = `import${type} { ${binding.exportedName} } from '${binding.specifier}';`;
  return statement + (binding.memberPath.length ? `\n// Access: ${accessName(binding)}` : '');
}

/**
 * Render a normalized example without altering its code bytes. A longer fence
 * keeps nested Markdown/backticks literal. Guides may reuse this pure helper.
 * @param {DocExample} example Source-owned example.
 * @param {number} [level=4] Optional title heading level.
 * @returns {string} Markdown ending in one newline.
 */
function renderDocumentationExample(example, level = 4) {
  return (example.title ? '#'.repeat(level) + ' ' + example.title + '\n\n' : '') + codeBlock(example.code, example.language) + '\n';
}

/** Rewrite repository-relative README links for the published npm package. */
function renderPackageReadme(source, config = defaultConfig) {
  const body = transformMarkdown(source.replace(/\r\n?/g, '\n'), {
    link: target => /^(?:[a-z][a-z0-9+.-]*:|#|\/)/i.test(target) || target === './LICENSE' || target === 'LICENSE'
      ? target : repository(config).blob + target.replace(/^\.\//, ''),
  });
  return finish('<!-- Generated from the root README.md by scripts/documentation.cjs. -->\n\n' + body);
}

function modelIndex(model) {
  const entities = new Map(model.graph.entities.map(entity => [entity.key, entity]));
  const pages = new Map();
  for (const page of model.pages) for (const key of page.entityKeys) pages.set(key, page);
  const bindings = key => model.graph.bindings.filter(binding => binding.entityKey === key)
    .sort((a, b) => compare(accessName(a), accessName(b)) || compare(a.specifier, b.specifier));
  function link(key, from, label, member) {
    const entity = entities.get(key);
    const page = pages.get(key);
    const name = label || entity?.qualifiedName || key;
    return page ? `[${name}${member ? '.' + member : ''}](${relativeTarget(from, page.route, page.anchors[key])})` : name + (member ? '.' + member : '');
  }
  return { entities, pages, bindings, link };
}

/** Project prose fields with their own source before combining contracts. */
function resolveContracts(contracts, model, destination) {
  const cache = new Map();
  return key => {
    const contract = contracts[key];
    if (!contract) return undefined;
    const upstream = contract.inheritedFrom && model.contracts[contract.inheritedFrom];
    // inheritedFrom also records local overrides' upstream relationship. Only
    // identical authored provenance establishes that this prose was borrowed.
    const borrowed = upstream && contract.source.file === upstream.source.file && contract.source.line === upstream.source.line;
    const ownerKey = borrowed ? contract.inheritedFrom : key;
    const binding = model.graph.bindings.find(item => bindingContractKey(item) === ownerKey);
    const entity = model.graph.entities.find(item => item.key === ownerKey || ownerKey.startsWith(item.key + '::'));
    const origin = model.pages.find(page => page.entityKeys.includes(binding?.entityKey || entity?.key))?.route || destination;
    const cacheKey = JSON.stringify([key, destination, origin]);
    if (cache.has(cacheKey)) return cache.get(cacheKey);
    const prose = value => value ? resolveMarkdown(value, model, contract.source.file, destination, undefined, origin) : value;
    const resolved = { ...contract, summary: prose(contract.summary),
      useWhen: contract.useWhen.map(prose), avoidWhen: contract.avoidWhen.map(prose),
      remarks: contract.remarks.map(remark => ({ ...remark, title: prose(remark.title), body: prose(remark.body) })),
      params: contract.params.map(item => ({ ...item, description: prose(item.description) })),
      typeParams: contract.typeParams.map(item => ({ ...item, description: prose(item.description) })),
      returns: prose(contract.returns), throws: contract.throws.map(prose),
      defaultValue: prose(contract.defaultValue), deprecated: prose(contract.deprecated), see: contract.see.map(prose),
      examples: contract.examples.map(example => ({ ...example, title: example.title
        ? resolveMarkdown(example.title, model, (example.source?.file || contract.source.file), destination, undefined, origin) : example.title })),
    };
    cache.set(cacheKey, resolved);
    return resolved;
  };
}

function renderContract(contract, heading, options = {}) {
  if (!contract) return '';
  const parts = [];
  const section = (title, body) => { if (body) parts.push(heading(title), body); };
  if (options.summary && contract.summary) parts.push(contract.summary);
  if (options.editorial !== false) {
    section('Deprecated', contract.deprecated);
    section('Use when', contract.useWhen.map(item => '- ' + item).join('\n'));
    section('Avoid when', contract.avoidWhen.map(item => '- ' + item).join('\n'));
    for (const remark of contract.remarks) {
      if (remark.title) parts.push(heading(remark.title));
      parts.push(remark.body);
    }
    section('Default', contract.defaultValue);
    section('Throws', contract.throws.map(item => '- ' + item).join('\n'));
    if (contract.examples.length) {
      parts.push(heading('Examples'));
      for (const example of contract.examples) {
        if (example.title) parts.push(heading(example.title, 4));
        parts.push(codeBlock(example.code, example.language));
      }
    }
    section('See also', contract.see.map(item => '- ' + item).join('\n'));
  }
  if (options.signature) {
    section('Type parameters', contract.typeParams.map(item => `- ${codeSpan(item.name)}: ${item.description}`).join('\n'));
    section('Parameters', contract.params.map(item => `- ${codeSpan(item.name)}: ${item.description}`).join('\n'));
    section('Returns', contract.returns);
  }
  return parts.filter(Boolean).join('\n\n');
}

/**
 * Render API references exclusively from the canonical graph/page/contract
 * metadata. Branch tables remain separate; support/external types never gain
 * invented imports. This function performs no file reads or writes.
 * @param {DocumentationModel} model Canonical source model.
 * @param {Config} config Versioned URL policy.
 * @returns {ArtifactSet} Relative POSIX routes to UTF-8/LF Markdown.
 */
function renderApiPages(model, config = defaultConfig) {
  const index = modelIndex(model);
  const contractKeys = Object.keys(model.contracts).sort(compare);
  const artifacts = {};
  for (const page of [...model.pages].sort((a, b) => compare(a.route, b.route))) {
    const contractAt = resolveContracts(model.contracts, model, page.route);
    const bindingContractAt = resolveContracts(model.bindingContracts, model, page.route);
    const publicHeadings = new Map();
    function reserve(anchor, key, name) {
      const previous = publicHeadings.get(anchor);
      if (previous && previous.key !== key) throw new Error(`API anchor collision on ${page.route}: ${anchor}`);
      if (!previous) publicHeadings.set(anchor, { key, name });
    }
    for (const key of page.entityKeys) {
      const entity = index.entities.get(key);
      if (!entity) throw new Error(`Missing API entity ${key}`);
      const bindings = index.bindings(key);
      for (const binding of bindings) {
        const name = apiHeadingName(page, entity, accessName(binding));
        const anchor = page.bindingAnchors[bindingContractKey(binding)] || markdownHeadingSlug(name);
        if (anchor !== markdownHeadingSlug(name)) throw new Error(`Non-natural API anchor ${anchor} for ${name}`);
        reserve(anchor, key, name);
      }
      const canonical = page.anchors[key];
      const primary = bindings.find(binding => markdownHeadingSlug(apiHeadingName(page, entity, accessName(binding))) === canonical);
      const name = apiHeadingName(page, entity, primary ? accessName(primary) : entity.qualifiedName);
      if (markdownHeadingSlug(name) !== canonical) throw new Error(`No natural heading for API anchor ${canonical}`);
      reserve(canonical, key, name);
      // A shorter alias must not erase the declaration's previously valid
      // natural anchor. The current public binding anchors remain authoritative.
      const declaredName = apiHeadingName(page, entity, entity.qualifiedName);
      const declared = markdownHeadingSlug(declaredName);
      if (!publicHeadings.has(declared)) reserve(declared, key, declaredName);
    }
    const used = new Set();
    function heading(title, level = 3) {
      let text = title;
      let counter = 1;
      while (used.has(markdownHeadingAnchor(text)) || publicHeadings.has(markdownHeadingAnchor(text))) text = `${title} (${++counter})`;
      used.add(markdownHeadingAnchor(text));
      return '#'.repeat(level) + ' ' + text;
    }
    const parts = [heading(page.title + ' API reference', 1)];
    const introductions = new Set();
    for (const { key, contract, full } of getPageIntroductions(model, page)) {
      if (full) {
        const resolved = contractAt(key);
        const identity = contractIdentity(resolved);
        if (!introductions.has(identity)) parts.push(renderContract(resolved, heading, { summary: true, signature: true }));
        introductions.add(identity);
      } else parts.push(resolveMarkdown(contract.summary, model, contract.source.file, page.route, undefined, page.route));
    }

    function memberContract(ownerKey, member) {
      const key = memberContractKey(ownerKey, member.name);
      const variantKey = contractKeys.find(candidate => candidate.startsWith(key + '::variant:') &&
        contractAt(candidate).source.file === (member.docSource || member.source).file &&
        contractAt(candidate).source.line === (member.docSource || member.source).line);
      return contractAt(variantKey || key);
    }
    function retainedOwner(entityKey, member, contract) {
      const owner = index.entities.get(member.ownerKey);
      // Only a direct, rendered declaring owner can terminate a sharing link.
      // Unrendered/branch-only owners and inherited chains fail closed.
      if (!owner || owner.key === entityKey || !index.pages.has(owner.key) ||
        contract.inheritedFrom !== memberContractKey(owner.key, member.name)) return undefined;
      const original = owner.members.find(item => item.name === member.name && sameSource(item.source, member.source));
      if (!original || original.ownerKey !== owner.key) return undefined;
      const upstream = memberContract(owner.key, original);
      if (!hasMemberDetail(upstream) || upstream.inheritedFrom || !sameSource(contract.source, upstream.source)) return undefined;
      return { contract: upstream, link: index.link(owner.key, page.route, undefined, member.name) };
    }
    function members(ownerKey, values, label, entityKey) {
      if (!values.length) return;
      const rows = ['| Member | Type | Requirement | Default | Description |', '| --- | --- | --- | --- | --- |'];
      for (const member of values) {
        const contract = memberContract(ownerKey, member);
        const inheritedOwner = contract?.inheritedFrom?.split('::member:')[0];
        const inherited = inheritedOwner ? `Inherited from ${index.link(inheritedOwner, page.route, undefined, member.name)}.` : '';
        const type = member.forbidden ? 'never' : member.type.text;
        const requirement = (member.forbidden ? 'Forbidden' : member.optional ? 'Optional' : 'Required') + (member.readonly ? ', readonly' : '');
        const defaultValue = contract?.defaultValue ?? (inheritedOwner ? `See ${index.link(inheritedOwner, page.route, undefined, member.name)}` : '—');
        rows.push(`| ${codeSpan(label + '.' + member.name)} | ${cell(codeSpan(type))} | ${requirement} | ${cell(defaultValue)} | ${cell([contract?.summary, inherited].filter(Boolean).join(' '))} |`);
      }
      parts.push(rows.join('\n'));
      // Complex member contracts stay in prose: table escaping must never
      // rewrite fenced examples or collapse authored remarks/constraints.
      for (const member of values) {
        const contract = memberContract(ownerKey, member);
        if (!hasMemberDetail(contract)) continue;
        parts.push(heading(label + '.' + member.name, 4));
        const owner = retainedOwner(entityKey, member, contract);
        if (owner && sameOwnedExamples(contract, owner.contract) && contractIdentity(contract) === contractIdentity(owner.contract)) {
          parts.push(`Contract: see ${owner.link}.`);
        } else if (owner && contract.examples.length && sameOwnedExamples(contract, owner.contract)) {
          // Inherited normalization clears defaults/remarks. Keep every other
          // local field verbatim rather than treating absence as a wildcard.
          parts.push(renderContract({ ...contract, examples: [] }, title => heading(title, 5), { signature: true }), `Examples: see ${owner.link}.`);
        } else parts.push(renderContract(contract, title => heading(title, 5), { signature: true }));
      }
    }
    function shape(ownerKey, value, label, entityKey) {
      members(ownerKey, value.members, label, entityKey);
      for (const [position, variant] of value.variants.entries()) {
        parts.push(heading(`${label} branch ${position + 1}`, 4), codeBlock(variant.text));
        members(ownerKey, variant.members, label, entityKey);
      }
    }
    for (const key of page.entityKeys) {
      const entity = index.entities.get(key);
      const canonical = page.anchors[key];
      const name = publicHeadings.get(canonical).name;
      used.add(canonical);
      parts.push('## ' + publicHeading(name));
      const contract = contractAt(key);
      if (contract?.summary) parts.push(contract.summary);
      const bindings = index.bindings(key);
      if (bindings.length) {
        parts.push(heading('Imports'));
        parts.push(codeBlock([...new Set(bindings.map(binding => legalImport(binding, entity)))].join('\n')));
      } else parts.push('Source-owned support type; no public named import.');
      const editorial = renderContract(contract, heading, { signature: !entity.signatures.length });
      if (editorial) parts.push(editorial);
      for (const binding of bindings.filter(binding => markdownHeadingSlug(apiHeadingName(page, entity, accessName(binding))) === canonical)) {
        const local = bindingContractAt(bindingContractKey(binding));
        if (local) parts.push(renderContract(local, heading, { summary: true, signature: true }));
      }
      if (entity.typeExpression || entity.typeParameters.length || entity.heritage.length || entity.kind === 'value' && !entity.signatures.length && entity.type) {
        parts.push(heading('Type'));
        const parameters = entity.typeParameters.length ? '<' + entity.typeParameters.map(parameter => parameter.name + (parameter.constraint ? ' extends ' + parameter.constraint : '') + (parameter.default ? ' = ' + parameter.default : '')).join(', ') + '>' : '';
        parts.push(codeBlock(entity.typeExpression ? `type ${entity.qualifiedName}${parameters} = ${entity.typeExpression}`
          : entity.kind === 'value' ? `${entity.qualifiedName}: ${entity.type.text}`
            : `${entity.kind} ${entity.qualifiedName}${parameters}${entity.heritage.length ? ' extends ' + entity.heritage.map(base => base.text).join(', ') : ''}`));
      }
      if (entity.signatures.length) {
        parts.push(heading('Signatures'));
        for (const signature of entity.signatures) {
          parts.push(codeBlock(signature.text));
          const own = contractAt(signature.key);
          const lexical = new Set(signature.sourceOwned ? signature.lexicalTypeParameters.map(parameter => parameter.name) : []);
          const current = new Set([...entity.typeParameters, ...signature.typeParameters].map(parameter => parameter.name));
          const declarationParameters = own?.typeParams.filter(parameter => lexical.has(parameter.name) && !current.has(parameter.name)) || [];
          const effective = declarationParameters.length ? { ...own, typeParams: own.typeParams.filter(parameter => !declarationParameters.includes(parameter)) } : own;
          const text = renderContract(effective, title => heading(title, 4), { signature: true, editorial: false, summary: own?.summary !== contract?.summary });
          if (text) parts.push(text);
          if (declarationParameters.length) {
            parts.push(heading('Source declaration type parameters', 4),
              'These describe the source declaration; the callable uses the instantiated signature above.',
              declarationParameters.map(parameter => `- ${codeSpan(parameter.name)}: ${parameter.description}`).join('\n'));
          }
          if (editorialIdentity(own) !== editorialIdentity(contract)) {
            const editorial = renderContract(own, title => heading(title, 4));
            if (editorial) parts.push(editorial);
          }
          if (signature.returns.members.length || signature.returns.variants.length) {
            parts.push(heading('Return shape', 4));
            shape(signature.key + '::returns', signature.returns, name + ' return', key);
          }
        }
      }
      if (entity.members.length || entity.variants.length) {
        parts.push(heading('Members'));
        shape(key, entity, name, key);
      }
      if (entity.references.length) {
        parts.push(heading('Related types'));
        for (const related of entity.references) {
          const referenced = index.entities.get(related);
          if (!referenced) continue;
          parts.push('- ' + index.link(related, page.route) + (referenced.visibility === 'external' ? ` — external dependency (${codeSpan(referenced.source.file)}); its upstream contract applies.` : referenced.visibility === 'support' ? ' — source-owned support type.' : ''));
        }
      }
      for (const [anchor, alias] of publicHeadings) {
        if (alias.key !== key || anchor === canonical) continue;
        used.add(anchor);
        parts.push('## ' + publicHeading(alias.name), `Alias/access for ${index.link(key, page.route, name)}.`);
        const aliasContracts = [...new Map(bindings.filter(binding => markdownHeadingSlug(apiHeadingName(page, entity, accessName(binding))) === anchor)
          .map(binding => [JSON.stringify(bindingContractAt(bindingContractKey(binding))), bindingContractAt(bindingContractKey(binding))] )).values()].filter(Boolean);
        for (const local of aliasContracts) parts.push(renderContract(local, heading, { summary: true, signature: true }));
      }
      parts.push(heading('Source'), `[${entity.source.file}:${entity.source.line}](${repository(config).blob}${entity.source.file}#L${entity.source.line})`);
    }
    artifacts[page.route] = finish(parts.filter(Boolean).join('\n\n'));
  }
  return artifacts;
}

/** Alphabetical public bindings; route deduplication must never remove rows. */
function renderCatalog(model, config = defaultConfig) {
  const index = modelIndex(model);
  const groups = new Map();
  const names = new Map();
  for (const binding of [...model.graph.bindings].sort((a, b) => compare(accessName(a), accessName(b)) || compare(a.specifier, b.specifier))) {
    const name = accessName(binding);
    if (!names.has(name)) names.set(name, new Set());
    names.get(name).add(binding.entityKey);
    const key = binding.entityKey + '#' + name;
    if (!groups.has(key)) groups.set(key, { name, bindings: [] });
    groups.get(key).bindings.push(binding);
  }
  const rows = [...groups.values()].sort((a, b) => compare(a.name, b.name) || compare(a.bindings[0].specifier, b.bindings[0].specifier)).map(group => {
    const binding = group.bindings[0];
    const entity = index.entities.get(binding.entityKey);
    const page = index.pages.get(binding.entityKey);
    if (!page) throw new Error(`Public binding ${bindingContractKey(binding)} has no API page`);
    const name = names.get(group.name).size > 1 ? binding.specifier + '#' + group.name : group.name;
    const target = relativeTarget('docs/api/INDEX.md', page.route, page.bindingAnchors[bindingContractKey(binding)] || page.anchors[entity.key]);
    const contract = model.bindingContracts[bindingContractKey(binding)] || model.contracts[entity.key];
    const summary = contract?.summary ? resolveMarkdown(contract.summary, model, contract.source.file, 'docs/api/INDEX.md', undefined, page.route) : '';
    return `| [${codeSpan(name)}](${target}) | ${entity.kind} | ${group.bindings.map(item => cell(codeSpan(legalImport(item, entity)))).join('<br>')} | ${cell(summary)} |`;
  });
  const navigation = [...new Map([...model.pages].sort((a, b) => compare(a.route, b.route)).map(page => [page.route, page])).values()]
    .map(page => `- [${page.title}](${relativeTarget('docs/api/INDEX.md', page.route)})`).join('\n');
  return finish('# API catalog\n\n## References\n\n' + navigation + '\n\n## Alphabetical symbol catalog\n\n| Public name | Kind | Legal imports and access | Description |\n| --- | --- | --- | --- |\n' + rows.join('\n'));
}

/**
 * Render the single LLM navigation export from supplied public page metadata.
 * Only the two contributor indices appear in Optional; internal/source/fixture
 * resources are excluded even when mistakenly supplied as page candidates.
 * @param {PublicPage[]} pages Actual discovered public pages.
 * @param {{name?: string, description?: string, version?: string}} packageMetadata Package facts.
 * @param {Config} config Versioned raw URL policy.
 * @returns {string} Deterministic llms.txt content, with raw page URLs.
 */
function renderLlms(pages, packageMetadata, config = defaultConfig) {
  const optional = new Set(['docs/development/INDEX.md', 'docs/development/sharepoint-validation/INDEX.md']);
  const visible = pages.filter(page => /^docs\/[^?#\\]+\.md$/.test(page.route) &&
    !/(?:^|\/)(?:\.docs|_sources|AGENTS\.md|fixtures?)(?:\/|$)/i.test(page.route) &&
    !page.route.split('/').some(part => part === '..' || part === '.') &&
    (!page.route.startsWith('docs/development/') || optional.has(page.route)));
  const unique = [...new Map([...visible].sort((a, b) => compare(a.route, b.route) || compare(a.title, b.title) || compare(a.summary, b.summary) || compare(a.area, b.area)).map(page => [page.route, page])).values()];
  const entries = list => list.map(page => `- [${page.title}](${repository(config).raw}${page.route})${page.summary ? ': ' + page.summary.replace(/\s+/g, ' ').trim() : ''}`).join('\n');
  const groups = new Map();
  for (const page of unique.filter(page => !optional.has(page.route))) {
    const area = page.area || 'Documentation';
    if (!groups.has(area)) groups.set(area, []);
    groups.get(area).push(page);
  }
  return finish('# ' + (packageMetadata.name || 'Package documentation') + '\n\n' + (packageMetadata.description ? '> ' + packageMetadata.description + '\n\n' : '') +
    (packageMetadata.version ? 'Version: ' + packageMetadata.version + '\n\n' : '') +
    [...groups].sort(([a], [b]) => compare(a, b)).map(([area, pages]) => '## ' + area + '\n\n' + entries(pages)).join('\n\n') +
    '\n\n## Optional\n\n' + entries(unique.filter(page => optional.has(page.route))));
}

module.exports = { renderApiPages, renderCatalog, renderPackageReadme, renderLlms, renderDocumentationExample };
