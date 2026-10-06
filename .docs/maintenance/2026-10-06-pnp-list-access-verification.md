# PnP list selector verification

Recorded 2026-10-06, Europe/Rome. Implementation remained on `dev` in the current checkout, based on HEAD `24eb52b195992d93564d50adf6f5e2a3ec60ea04`. Verification took place before commit and push, which the user subsequently authorized. Pre-existing `AGENTS.md` was preserved by SHA-256 comparison. No dependency, package version, peer, entry-point, package allowlist or historical API/package fixture changes were made.

## Implemented scope

- One `createSPFxPnPListService` accepts exact title strings or the readonly title/id/url/path selector. Every operation uses the actual ordinary/batched client. Path selection derives the configured explicit web without metadata HTTP.
- The historical title hook/types remain unchanged; `ById`, `ByUrl` and `ByPath` share its extracted lifecycle implementation.
- Public documentation and the existing test WebPart PnP panel/registry cover all modes. The sample has actual concrete hooks, explicit read actions and accessible inputs; the original title CRUD demo is unchanged.
- Compatibility approvals are narrowly scoped and preserve the baseline. Isolated consumer probes cover old/new typed calls, generic returns, root/deep imports and malformed selector types.

## Final local commands

The controller ran the following on the corrected source, not only the worker reports:

Pre-commit staging also checked newly added files. Fifteen whitespace-only lines inherited by the extracted hook were normalized; executable lines and runtime behavior were unchanged.

| Command/observation | Result |
| --- | --- |
| `npm run build:library` | PASS, exit 0 |
| `npm run build:app`, after library build | PASS, exit 0; TypeScript/lint/webpack tasks complete |
| `npm run verify` | PASS, exit 0; 460 tests, 0 failures/skips/cancellations; both workspace typecheck/lint; example inventory 44 hooks/4 providers; runtime/docs checks; 82 historical declaration modules |
| `npm run verify:package` | Wrapper exit 0; installed a real external 367-file / 196469-byte tarball consumer; no workspace symlink; public/deep TypeScript probes, lint, webpack and solution generation completed. Clean ship output remains UNVERIFIED as explained below. |
| `git diff --check` | PASS |
| Historical fixtures/package metadata/lockfile diff | Empty |
| Pre-existing `AGENTS.md` hash | Unchanged |

Current-session raw logs are `/private/tmp/pnp-list-final-corrected-verify.log` and `/private/tmp/pnp-list-final-corrected-package.log`. Temporary external consumers were cleaned by the existing verifier. These temporary paths are supporting diagnostics; the durable results and limitations are recorded here.

## Ship-toolchain limitation

The package subprocess returns zero after completing its tasks, but emits `The build failed because a task wrote output to stderr` and `Exiting with exit code: 1`. Its preceding stderr consists of old Baseline browser mapping/Browserslist metadata notices, not feature lint/type errors. Do not use the verifier's blanket ship-success message as a clean ship-readiness claim.

Read-only tracing established that locked `baseline-browser-mapping` 2.8.20 emits its stale-data notice using an unconditional date comparison, without a supported suppression flag. Browserslist has an old-data environment guard, but it cannot suppress that separate notice. `@microsoft/gulp-core-build/lib/logging.js` records stderr in ship mode and calls asynchronous `exitProcess(1)` inside a process exit listener; that callback cannot reliably alter the already-exiting process status. No dependencies, stderr filtering, warning policy or toolchain configuration were changed. The clean ship check is therefore explicitly UNVERIFIED, although package content/type/webpack/solution evidence above is available.

## Authenticated SharePoint read QA

Browser: connected Microsoft Edge, existing authenticated test tenant. An independent workbench tab mounted the local `SPFxReactToolkitTest` sample. The library was rebuilt and the repository's existing local Gulp serve process restarted on HTTPS localhost:4321; the final sample was reloaded after its copy correction. The server remains available for continued local testing. Site/tenant identifiers are omitted from this versioned record.

Flow: workbench -> PnPjs -> explicit title/GUID/URL/path query -> first 10 items of the same existing Documents library. The displayed inventory contained 49 covered symbols, including the three new hooks. No tenant data writes, permission changes, publication or deployment were performed.

| Check | Observation | Status |
| --- | --- | --- |
| Meaningful page and controls | Actual workbench/sample loaded; no framework error overlay; distinct inputs/buttons for GUID, URL and path | PASS |
| Initial blank modes | New actions disabled; no populated list results. Editing values displayed neutral `No items loaded.` before requesting reads | PASS |
| Title, canonical GUID, decoded server-relative root and decoded web-relative root | All four returned 10 items with identical descending IDs: 138, 137, 136, 132, 131, 130, 129, 128, 127, 126 | PASS |
| Root path with spaces | URL/path used the decoded Shared Documents root and resolved the same library | PASS |
| Nullable Title | New demos displayed `(untitled)` for actual empty optional Title values | PASS |
| Invalid GUID | Query exposed the local validation message; no framework failure or empty-success text; other modes retained their results | PASS |
| Recovery/normalization | Uppercase GUID enclosed in braces with surrounding whitespace recovered and returned the same first page | PASS |
| Separate hook state | GUID failure/recovery did not clear title, URL or path result sets | PASS |
| Console | No query-flow unhandled rejection observed in the final interactions. Console is not globally clean: pre-existing/host icon-registration notices and an earlier SharePoint mysite-cache promise error were recorded | LIMITED |
| Tenant mutations/batch multipart, permission-denied cases, root/subweb/cross-site and literal `%`/`#` server behavior | Not exercised by this read-only QA; local transport/lifecycle tests remain separate evidence | NOT EXECUTED |

Visual inspection showed readable controls and successful result rows. The current-session native screenshot is `/private/tmp/pnp-list-access-qa.jpg`; it captures the URL/path query controls and identical result pages. The source checklist remains a reusable template; this table records the actual run's narrower coverage.

## Independent review

Service, hook, compatibility, public-doc and sample tasks each passed independent scoped reviews. The final whole-feature reviewer found an Important prose mismatch: partial item batch failures set hook error and resolve successful IDs/undefined rather than rejecting. It also found a Minor pre-query empty-message issue. A single correction wave fixed the batch/getById failure guidance and changed only the new sample empty string to `No items loaded.`. The scoped correction re-review marked both findings ADDRESSED, with no new material issues.

Historical runtime semantics were preserved: hook `getById` service failures resolve undefined with error, while missing context rejects; hook partial batches publish summary error and resolve their established values, while service rejection still rejects batch promises. No remaining feature-code review finding is open. Clean ship output and the unexecuted tenant cases above remain explicit verification limits.
