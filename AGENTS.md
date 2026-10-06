# Agent instructions for the monorepo

These instructions apply to the entire repository, including `packages/`, `apps/`, `tests/`, `scripts/`, `docs/`, and `.docs/`, and to subagents assigned to the work.

## Checkout and worktrees

- Work in the current repository checkout. Do not create or use worktrees automatically, even when a skill recommends Git isolation.
- Use a worktree only when the user explicitly requests one. Permission to create a branch does not authorize creating a worktree.
- If a session starts in a worktree, tell the user and agree on using the main checkout before modifying files. Do not create additional worktrees.

## Choose a branch before making changes

1. At the start of a new assignment that requires changes, identify the current branch and Git status with `git branch --show-current` and `git status --short`.
2. Before modifying files or creating/switching branches, ask the user:
   > Would you like to work on the current branch `<branch>`, or create a new branch in the same checkout?
3. Wait for the user's choice. You may perform read-only inspections while waiting, without modifying the checkout.
4. If the user has already explicitly selected a branch for that assignment, follow that choice without asking again. The choice applies to continuations and subagents working on the same assignment.
5. If the current branch is selected, stay on it. If a new branch is selected, create it in the same checkout from the current HEAD. Use the user-provided name or a relevant name with the `apvee/` prefix.
6. If HEAD is detached, ask which branch to use before making changes.

## Existing changes and coordination

- Preserve pre-existing staged and unstaged changes and untracked files. Do not use reset, clean, stash, or file checkout to clear the branch without authorization.
- If switching branches encounters conflicts or could overwrite changes, stop and explain the problem to the user. Do not work around it by creating a worktree.
- The orchestrator coordinates branch selection. Subagents use the same checkout and must not create or switch branches independently.
- Choosing a branch does not authorize automatic pushes or merges. Perform them only when requested by the user.

## Documentation locations

- Keep public product documentation, API references, installation, development, and SPFx validation guides in `docs/`.
- Store internal plans, maintenance records, findings, review notes, and verification evidence in `.docs/maintenance/`.
- Store Superpowers plans and specifications in `.docs/superpowers/`; this directory contains local ignored planning material.
- Keep reusable test fixtures in `tests/fixtures/`, not in documentation or evidence directories.
- Keep this `AGENTS.md` at the repository root. Use these documentation locations even when a skill suggests a different default path.
