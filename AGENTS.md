# Agent instructions for the monorepo

These instructions apply to the entire repository, including `packages/`, `apps/`, `tests/`, `scripts/`, `docs/`, and `.docs/`, and to subagents assigned to the work.

## Instruction boundaries and working principles

- Apply these repository rules within applicable system/developer instructions, runtime capabilities, and permissions. Direct user instructions for the assignment take precedence over repository defaults.
- Repository rules take precedence over optional skill defaults. Follow mandatory runtime requirements for skill invocation, disclosure, communication, and delegation; the lightweight workflows below apply where those requirements leave discretion.
- Read applicable repository instructions before making changes. Read architecture and development documentation relevant to the task rather than every document for every edit.
- Use documentation to identify intended contracts and conventions; use code and tests to establish current behavior. Surface material discrepancies instead of silently treating either source as infallible.
- Prioritize correctness and maintainability. Optimize speed, token/context usage, and tool usage within those constraints, including the cost of likely rework.
- Use the lightest workflow that provides sufficient confidence. Complete authorized work without expanding its scope into unrelated refactoring or migration.

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

## Repository architecture and compatibility

- `packages/spfx-react-toolkit/src/` owns the public library; `apps/spfx-react-toolkit-test/` owns the private SPFx sample and host configuration. Root `tests/` and `scripts/` own behavioral tests and verification tooling.
- Consult `README.md` and `docs/DEVELOPMENT.md` for workspace boundaries, supported tooling, and commands. Consult `docs/SHAREPOINT-VALIDATION.md` when real-host validation is relevant. Read affected API documentation in `docs/` when changing public behavior.
- The sample consumes the library's compiled `lib` entry point. Build the library before the app; after library source changes, rebuild it and restart a running app serve process before evaluating the new behavior.
- Preserve public imports, exports, declaration contracts, package contents, peer dependency contracts, and provider instance isolation unless the user authorizes the relevant change.
- Follow repository React, TypeScript, Node, and SPFx compatibility constraints. Do not upgrade the toolchain or dependencies merely to simplify an implementation.
- Identify affected modules, direct dependencies, consumers, and lifecycle boundaries before changing production code. Follow approved migration plans and checkpoints when present.
- If documentation is missing, inspect relevant implementation and tests. Do not invent architectural requirements, documentation paths, or application/domain layers that the repository does not have.

## Proportional execution

Assess complexity, uncertainty, blast radius, and irreversibility internally. Start with the lightest plausible level and escalate when evidence reveals greater risk or coupling.

| Level | Scope | Sufficient workflow |
| --- | --- | --- |
| L0 — Mechanical | Verified non-behavioral change with negligible contract risk | Inspect, edit, relevant static verification |
| L1 — Bounded | Localized behavior change inside an existing implementation | Inspect, reason, implement, targeted regression/behavior verification |
| L2 — Feature | Bounded new capability within the existing architecture | Focused exploration, lightweight design, implement, meaningful tests, review when useful, verify |
| L3 — Architectural | Public contracts, persistence, migrations, infrastructure, or coupled subsystems | Explore, evaluate alternatives, design, plan, implement incrementally, test, review, verify |

- A rename, import fix, metadata edit, or behavior-preserving refactor is L0 only after checking that it does not affect public API, packaging, runtime behavior, or other consequential boundaries.
- Prefer the higher level when uncertainty concerns security, architecture, persistence, public contracts, destructive actions, or broad impact. Do not retain an undersized workflow merely because the task initially looked simple.
- Parallel execution is a separate decision, not a risk level. Large tasks need dependency-aware decomposition; they do not automatically require agents.
- Routine changes do not need permanent plans, specifications, or intermediate design approval unless applicable runtime requirements or user instructions require them.
- **Branch selection above remains mandatory for every assignment requiring changes, including L0.** Once selected, retain that choice for continuations and subagents.

## Autonomy and decision points

- After branch selection, proceed autonomously when intent is clear, the outcome is unambiguous, changes are reasonably reversible, and no consequential product/architecture decision or unexpected public contract change remains unresolved.
- Ask when materially different outcomes require user choice, required information cannot be discovered safely, or a destructive/irreversible action lacks authorization. Respect authorization already provided.
- Obtain discoverable answers from repository context, code, and tests instead of asking the user. Continue useful independent work while a required answer is pending.
- Honor explicit stop, review, and approval checkpoints. Complete all authorized preparation and verification before asking for final approval of an action that requires it.
- Branch selection and local implementation do not authorize publication, deployment, tenant permission changes, pushes, or Git merges. Follow explicit user authorization for those actions.

## Model and reasoning preferences

These are preferences for controls actually exposed by the runtime, not a claim that the current session can change its own model or reasoning effort.

| Role | Preferred model | Reasoning | Suitable work |
| --- | --- | --- | --- |
| Worker | GPT-6.1 Sol (`gpt-6.1-sol`) | low | Deterministic, low-risk, well-specified execution |
| Engineer | GPT-6.1 Sol (`gpt-6.1-sol`) | medium | Normal implementation, bounded debugging, tests, and repository investigation |
| Reasoner / Orchestrator | GPT-6.1 Sol (`gpt-6.1-sol`) | high | Consequential ambiguity, difficult debugging, architecture, decomposition, integration, or review |

- Use GPT-6.1 Sol for all roles where available.
- Model tier and reasoning effort are separate controls. Escalate for concrete uncertainty, failed attempts, hidden coupling, or higher impact; reduce effort when remaining work is deterministic, where supported.
- Use `xhigh` only when `high` is insufficient for a concrete consequential task. Do not use higher effort merely because it is available.
- Select only models/settings supported by the current runtime. If GPT-6.1 Sol is unavailable, explain the limitation and ask the user to select an alternative before invoking another model. Never claim a model or reasoning change that did not occur.
- Exclude GPT-6 Astra from routine selection, escalation, delegation, and review. Agent-initiated Astra use requires a concrete unmet requirement, an explanation of why GPT-6.1 Sol at an appropriate supported reasoning effort is insufficient, and explicit user confirmation.
- An explicit user selection of Astra counts as confirmation. If the runtime identifies the active session as Astra and no such authorization is established, pause and ask for confirmation or a non-Astra selection. Do not infer the active model from assumptions or branding.
- Do not create a subagent merely to change model or reasoning settings. Model preferences do not authorize delegation.

## Focused exploration and engineering

- Prefer `rg`/`rg --files`, then focused reads: task references, relevant implementation/tests, direct dependencies/consumers, and additional documentation as needed. Expand context to answer a concrete question; stop once evidence is sufficient to act safely.
- Avoid duplicate searches, unnecessary rereads, speculative repository-wide discovery, and tool calls made only to show activity. Batch independent read-only operations when supported; keep dependent operations and mutations ordered.
- Make the smallest coherent change that fully solves the task. Prefer existing conventions, explicit behavior, clear interfaces, focused modules, and minimal coupling.
- Avoid unrelated cleanup, speculative abstractions, premature generalization, unnecessary dependencies, and theoretical future-proofing. A shared abstraction needs demonstrated common semantics or a necessary shared contract; two consumers alone do not establish that need.
- Preserve readable code and meaningful error handling. Do not compress code or weaken types to reduce line count or token usage.
- In React, keep rendering pure, derive values instead of duplicating state, and use effects for external synchronization and cleanup. Separate reusable business logic from rendering where existing repository boundaries support it.
- In TypeScript, use precise types, useful discriminated unions, and validation at untrusted boundaries. Avoid `any`, unsafe casts, and non-null assertions unless their safety is established.

## Required documentation and sample coverage

- Every new public library capability, including generic React hooks, controls, providers, services and helpers, must include all three deliverables below before it is considered complete. Extend the same deliverables when adding behavior to an existing capability.
- **JSDoc:** Document the public API in source, including exported types and their public members, parameters/props, return values, behavior, relevant constraints and a usage example. Keep IntelliSense documentation consistent with the implementation.
- **Markdown:** Add a dedicated section in the relevant public documentation under `docs/`, covering usage, API shape, examples, dependencies and relevant limitations. Link it from the appropriate documentation index; a changelog entry alone is insufficient.
- **Test web part:** Add a working example or test scenario in `apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/` that exercises the capability and makes the expected behavior observable. A symbol listing or static description alone is insufficient. Update the demo registry and applicable coverage checks.
- The web part scenario complements the automated behavioral tests required by this document; it does not replace them. Report whether the scenario was actually executed, and distinguish local verification from authenticated SharePoint-host validation.

## Debugging and testing

- For an evident root cause, inspect, fix, and verify the regression. For uncertain causes, reproduce, gather evidence, form/test a hypothesis, fix the cause, and verify. Increase depth for failed fixes, distant symptoms, timing, concurrency, caching, or complex state transitions.
- Keep tests in the established root `tests/*.test.cjs` convention and reusable fixtures in `tests/fixtures/`. Reuse existing harnesses where appropriate. Do not introduce nearby `__tests__/` directories or reorganize tests merely because a generic workflow suggests them.
- Test externally observable behavior and meaningful boundaries, not implementation details. Match testing effort to behavioral risk.
- Mechanical or documentation-only changes normally need relevant static validation and inspection, not new behavioral tests.
- For bugs, add a meaningful regression test when practical: confirm it fails for the intended reason before the fix, then passes afterward. If this is impractical, use the strongest practical reproduction and verification, and disclose the limitation.
- For normal new behavior, add/update meaningful tests; TDD is optional unless required by applicable instructions or useful for the risk. Prefer strict test-first work when practical for security/authorization, financial calculations, parsers, persistence transformations, destructive migrations, and correctness-critical algorithms or state transitions.
- Once strict test-first work is selected, preserve the failing-test, minimum implementation, passing-test, and re-verification sequence.

## Verification and completion evidence

- Before claiming completion, a fix, passing checks, or readiness, identify the command/observation supporting that specific claim, run it fresh, inspect output and exit status, and confirm that it proves the claim.
- Use the smallest sufficient verification during iteration; this does not waive documented project gates. Do not rely solely on confidence, old results, code inspection, partial checks, or a subagent's success report.
- Root `npm test` runs Node behavioral tests. `npm run typecheck` and `npm run lint` cover the workspaces; `npm run build:library` precedes `npm run build:app`.
- Run `npm run verify` and `npm run verify:package` before requesting development review, as required by `docs/DEVELOPMENT.md`. These commands do not authorize publication or deployment.
- For instruction/documentation-only edits, inspect the diff, check whitespace and referenced paths/commands, and assess internal consistency. Run additional documentation checks if the affected content requires them; full application builds are not needed solely for an `AGENTS.md` edit.
- Build, local tests, and package verification do not establish authenticated SharePoint-host behavior or tenant permissions. Keep real-host results separate and label unavailable checks as not executed or blocked; do not present simulations as tenant passes.
- Confirm requested behavior, scope, likely regressions, justified complexity, and material limitations before completion. If a required check cannot run, report what was verified and what remains unverified.

## Skills, plans, and delegation

- Use installed specialist skills for concrete relevant triggers, honoring mandatory runtime requirements. Do not assume a skill/plugin exists or load optional workflows merely to increase process depth.
- Where discretionary, use brainstorming for consequential design ambiguity, structured debugging for uncertain causes, plans for meaningful sequencing/handoff, strict TDD for appropriate risks, and independent review where it materially improves confidence.
- Create durable plans, specifications, investigation records, or decision artifacts only for architectural memory, migration, handoff, multi-session work, or other lasting value. Use the documentation locations above.
- Delegate only when supported and authorized by the user or applicable instructions, and when independent workstreams or independent analysis justify coordination cost. Do not interpret task size or model routing as authorization.
- Give agents clear scope, evidence requirements, and ownership boundaries; avoid concurrent edits to the same files. The orchestrator owns branch selection and integration, and verifies agent output against the actual diff and relevant checks.
- Agent reports should contain only decision-relevant results, changed files, verification evidence, and remaining issues. Prefer `RESULT`, `CHANGES`, `EVIDENCE`, and `ISSUES` over chronological logs.

## Communication

- Communicate concisely without reducing reasoning depth, implementation quality, verification, or accurate error reporting. Apply routine rules without repeatedly restating them.
- Surface required decisions, consequential assumptions, blockers, material scope changes, and verification limitations. Follow runtime progress-update and skill-disclosure requirements; avoid narration of every routine read, edit, or tool call.
- Final responses should state the result, meaningful changes, verification, and relevant caveats in proportion to task complexity. Distinguish completed work from pending or unverified work; avoid chronological recaps.
