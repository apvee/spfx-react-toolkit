# Internal documentation and dependency ownership — 2026-10-06

The user approved moving internal documentation into .docs, then explicitly selected both Fluent packages as shared peer dependencies.

## Documentation layout

- Public API, installation, development and SPFx validation guides remain in docs/.
- Versioned maintenance records/evidence moved to .docs/maintenance/.
- Eight existing local Superpowers plans/specifications moved to .docs/superpowers/, retaining their ignored status. All original planning bytes were preserved.
- Root AGENTS.md remains English and defines these locations for future assignments, without changing its branch/worktree policy.
- Tests and immutable baseline fixtures stay in tests/fixtures. Current links and the public-doc verifier were updated; historical paths inside command logs/plans remain historical.

67 existing internal files were copied/moved without data loss; only current public-link targets required adjustment. The public-doc verifier now scans docs/ as public documentation without internal-directory exceptions.

## Fluent ownership selected by the user

| Package | Library | SPFx test app |
|---|---|---|
| @fluentui/react-migration-v8-v9 ^9.9.12 | Mandatory peer + development dependency | Runtime dependency supplying the peer |
| @fluentui/react-theme ^9.2.0 | Mandatory peer + development dependency | Runtime dependency supplying the peer |
| tslib 2.3.1 | No direct dependency: no own source/emitted helper imports | Supplied transitively by the SPFx runtime packages that declare it |

Both Fluent modules are actually imported by public theme helpers; the public declarations reference Theme. This is therefore a deliberate consumer contract change, not deletion of unused imports. The initial recommendation kept them as owned dependencies; the user's explicit sharing requirement superseded that choice. Existing SPFx/React/PnP peer ranges, package name/version/entry paths and all82 declaration contracts remain unchanged.

No existing package versions changed in the lockfile; changes are workspace dependency-role maps only. The immutable historical package fixture remains unchanged. verify-api permits precisely the user-approved Fluent peer move and removal of unused direct tslib, keeps strict map equality otherwise, and now checks emitted JS/declaration imports against declared dependencies/peers to prevent reliance on root hoisting or transitives.

## Verification and review

- Consumer before change: PASS with the original dependency ownership.
- Clean install/build,90 tests, typecheck/lint/public docs/runtime/examples and API82: PASS after the peer change.
- Real packed tarball installed outside the monorepo: public/deep imports, shared Fluent realpaths, TypeScript, ship bundle and solution PASS.
- Additional consumer started with no copied lockfile or node_modules and pre-existing host Fluent versions9.9.12/9.2.0: PASS. Independent reviewer confirmed the toolkit is not a workspace symlink and host/library resolve the exact same Fluent package paths.
- Fresh consumer initially failed typecheck before Sass declarations existed. Fixed validation order to bundle:ship→typecheck→package:solution. All checks remain; failure and retry are recorded.
- Independent docs_peer_integration_review (GPT-6.1 Sol) PASS: relocation/link integrity, immutable fixtures, lock-role changes, mandatory peers, import checks and clean-consumer bootstrap. No runtime source was modified.

Root server/script and earlier cleanup work are preserved. No worktree, branch switch, commit/push/publish/deploy performed. Tenant authentication remains NOT RUN/BLOCKED.
