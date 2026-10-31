# Maintain the documentation

[Documentation home](../INDEX.md) · [Development index](INDEX.md)

## Markdown formatting

Write each prose paragraph on a single source line and let the editor or renderer wrap it visually. Keep each list item's prose on one line, with separate lines for nested items or blocks. Preserve the line breaks required by code blocks, tables, headings and intentional Markdown hard breaks.

## Content ownership

| Content | Owner |
| --- | --- |
| Product introduction and minimal example | Root `README.md` |
| Installation, compatible peers and verified baseline | `docs/getting-started/installation.md` |
| First complete integration and provider choice | `docs/getting-started/` |
| Task-oriented selection and composition | `docs/guides/` |
| Detailed public contracts | The capability's page under `docs/api/` |
| Repository setup and verification commands | `docs/development/` |
| Repeatable authenticated host procedures | `docs/development/sharepoint-validation/` |
| Dated reviews, findings, run results and evidence | Local `.docs/maintenance/` |

Keep required reusable test inputs in versioned `tests/fixtures/`, never in local evidence directories. Existing historical `.docs` files tracked on older branches are preserved; new local material is ignored. Public guides must remain useful without access to that material.

## API reference pattern

Use one page for a capability with a substantial independent contract. Keep related types, compound parts and small coherent helper families together. An index should introduce choices and link to the reference, rather than repeat signatures, full examples or default-value tables.

Use the following order where applicable:

1. Purpose and a concrete use case.
2. Supported import and a minimal valid example.
3. Provider, shared runtime, client and permission requirements.
4. Signature, public types, parameters/props and verified defaults.
5. Return values, callbacks, errors and asynchronous ownership.
6. Limitations, compatibility and absence behavior.
7. Related guide, source and observable sample scenario.

Source JSDoc and the Markdown reference describe the same public behavior. JSDoc supports IntelliSense; guides explain selection and composition. A source symbol listing alone does not replace a contract or an observable working sample. Follow [AGENTS.md](../../AGENTS.md) for the required JSDoc, Markdown and test-web-part deliverables whenever adding or changing a public capability.

## Generated discovery pages and README

The root README is the editorial source for the npm README. The generator changes repository-relative documentation links to absolute GitHub links, preserves literal fenced examples and keeps the packaged LICENSE link local.

The alphabetical [API catalog](../api/INDEX.md) is generated from all clean package entry points, including root/domain exports, imported re-exports, aliases, public types and dedicated native Property Pane factories. Its descriptive sections also provide navigation by API area; individual reference folders link back to those sections instead of maintaining separate indexes. Keep the documentation home, API reference, development guide and SharePoint validation guide as the four entry points with distinct purposes. The [route map](../../scripts/documentation-routes.json) maps source modules to their reference page and optionally overrides an individual symbol. Keep the entry's import contract and reference accurate when changing that map. Unknown public source modules fail generation rather than disappearing from the catalog. Style namespaces link to the dedicated qualified style-member catalog.

After editing the root README, public exports or routes, run from the root:

```bash
node scripts/documentation.cjs
node scripts/documentation.cjs --check
npm run verify:public-docs
```

Commit both generated artifacts with their source changes. `--check` reports stale artifacts without writing them. The documentation verifier checks generation, every clean-entry export route, all React family members/JSDoc, native factories, helper/service inventories, qualified styles/JSDoc, local paths, Markdown fragments and supported imports in TypeScript/JavaScript fenced examples. Generated descriptions summarize source JSDoc; they are not independent proof of behavior or example correctness.

## Moves and current links

When splitting or moving a document, update current references to its canonical destination, including references in AGENTS.md, the route map and the README. Rebase source links and Markdown fragments from the document's original location. Use natural heading anchors and keep indexes focused on current content.

Remove superseded pages once their unique content has been moved. A reference has one canonical owner; indexes and task guides link to it. Historical section menus and artificial compatibility anchors are not maintained. Package import, saved-data and public API compatibility remain separate documented contracts.

For documentation-tooling changes, run the targeted checks:

```bash
node --test tests/documentation-tooling.test.cjs
npm run verify:public-docs
git diff --check
```

Review examples against actual exports and signatures. Link checks do not execute examples, verify remote URLs or establish authenticated SharePoint behavior.
