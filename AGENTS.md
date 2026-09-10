# AGENTS.md

## Purpose

This repository is the active development workspace for **WinTuner GUI**.

Codex must treat this file as the primary set of development rules for all work in this repository.

Project status, completed work, architectural context and roadmap are documented separately in:

`PROJECT_STATUS.md`

Read both files before making significant changes.

---

## Repository and Branch Rules

Repository:

`manuelhoefler17-gif/WinTuner-GUI`

Branch policy:

- `main` is the stable branch.
- `Test` is the active development branch.
- Do not develop directly on `main`.
- Do not commit directly to `main`.
- New work must be performed on `Test` unless explicitly instructed otherwise.
- Changes reach `main` through a Pull Request.
- After a PR is merged, synchronize `Test` with `main`.

Before making changes:

1. Check the current branch.
2. Check `git status`.
3. Confirm the working tree is clean or understand every existing local change.
4. Do not overwrite unrelated user changes.

Never assume the repository is clean.

---

## Local Paths

Primary development repository:

`C:\Intune\WinTuner\Entwicklung`

Do not modify this separate production/local copy unless explicitly instructed:

`C:\Intune\WinTuner\WinTuner_GUI.ps1`

All normal development work must remain inside the repository.

---

## PowerShell Requirements

Use **PowerShell 7**.

Preferred executable:

`pwsh.exe`

Do not use Windows PowerShell 5.1 for normal development or testing unless specifically required for a compatibility test.

The GUI should normally be started from the repository with:

`pwsh.exe -File .\WinTuner_GUI.ps1`

---

## Development Philosophy

Prefer:

- small changes
- incremental changes
- explicit validation
- conservative behavior
- clear error handling
- safe failure modes
- tests before commits

Avoid:

- broad rewrites without need
- unrelated cleanup during feature work
- speculative refactoring
- large multi-feature commits
- silent behavior changes
- hidden fallback behavior

Preserve stable working behavior unless a change is intentionally requested.

---

## Editing Rules

Before changing a file:

- inspect the relevant code first
- understand surrounding state and dependencies
- identify all related callers when changing shared behavior

For scripted or automated edits:

- use `$ErrorActionPreference = 'Stop'`
- verify expected matches before replacing text
- require unique matches when appropriate
- modify content in memory first
- validate PowerShell syntax before writing
- only write after validation succeeds

Prefer:

- AST-aware changes for significant PowerShell functions
- complete known-block replacement over fragile partial replacements
- line-based scoped edits over broad regex replacements

Avoid:

- uncontrolled global replacements
- placeholder code such as `...`
- hand-wavy partial patches
- edits outside the intended file set

---

## PowerShell Syntax Validation

For PowerShell files, validate syntax after meaningful edits.

Use parser-based validation where practical.

A change must not be considered complete if syntax errors remain.

---

## Git Safety Rules

Do not use:

`git add .`

Do not use:

`git add -A`

Stage only explicitly intended files.

Examples:

`git add .\WinTuner_GUI.ps1`

`git add .\README.md`

Before committing:

1. inspect `git diff`
2. run `git diff --check`
3. verify only intended files changed
4. run relevant tests
5. confirm the affected workflow works

Do not commit unless explicitly instructed or clearly authorized.

Do not push unless explicitly instructed or clearly authorized.

Do not create or merge Pull Requests without explicit authorization.

---

## Commit Style

Use concise Conventional Commit-style messages where practical.

Examples:

`fix: validate exact IntuneWin package before upload`

`feat: show discovery cache usage in status`

`refactor: centralize package upload state`

`docs: refresh README for 0.10.14`

Keep each commit focused on one logical change.

---

## GUI State Safety

WinTuner GUI contains stateful WinForms behavior.

When changing UI state logic:

- identify all locations that can mutate the same state
- avoid blindly setting controls to enabled
- prefer central state calculation
- preserve safe disabled states on validation failure
- test login/logout transitions
- test selection changes
- test version changes
- test success and failure paths

Do not assume a button is safe to enable merely because a previous operation succeeded.

---

## Package Build and Upload Safety

Package upload safety is critical.

Do not trust a package only because a directory exists.

Upload readiness must consider:

- tenant connection
- selected package
- selected version
- built version
- package metadata
- expected `.intunewin`

`win32LobApp.json` is used to determine the exact expected `.intunewin` filename.

Do not replace exact package validation with broad wildcard checks.

Keep click-time validation as a final safety layer even if UI state already indicates readiness.

---

## Package Reuse

Package reuse supports GUI restarts only after centralized artifact validation succeeds.

Existing package folders from a previous GUI session must never be trusted solely because they exist.

Cross-session reuse must validate the selected package ID and version, readable metadata, matching metadata version, a safe .intunewin leaf filename, and the exact non-empty non-reparse-point file before marking a package as reusable.

Refer to `PROJECT_STATUS.md` for current roadmap decisions.

---

## Self-Update Safety

The self-update flow is security- and reliability-sensitive.

Preserve:

- temporary download before replacement
- version validation
- optional SHA256 validation
- backup before replacement
- rollback on failed replacement/restart
- automatic restart after success
- standalone release bootstrap
- development checkout protection

A Git development checkout must not silently overwrite local `Modules` or `Workers` files with release files.

If required development files are missing, fail clearly rather than mixing release and development files.

---

## Discovery Safety and Performance

The discovery workflow may process large Intune application inventories.

Preserve:

- Graph pagination
- retry handling
- detected-app caching
- persistent WinGet discovery caching
- isolated worker processes
- batching
- cancellation
- duplicate handling

Avoid changes that reintroduce large repeated WinGet lookups or UI blocking without a clear reason.

When changing discovery behavior, test both:

- a fresh run
- a cache-heavy repeated run

---

## Logging

Important state changes, warnings and failures should be logged.

Logs should help explain:

- what operation was attempted
- what package/version was involved
- why an operation was blocked
- whether cached or fresh data was used
- whether rollback or recovery occurred

Avoid logging secrets, tokens, passwords or sensitive authentication material.

---

## Documentation

When behavior changes materially, consider updating:

- `README.md`
- `CHANGELOG.md`
- `PROJECT_STATUS.md`

Do not document features that are not implemented.

Keep documentation aligned with actual behavior.

---

## Release Discipline

Do not create a new release casually.

Before a release:

- complete the changelog
- verify version references
- run final E2E tests
- verify syntax/tests
- verify standalone release behavior
- verify release assets
- generate checksum assets where required

Do not tag or publish a release unless explicitly instructed.

---

## Testing Expectations

For every meaningful change, test the affected path.

Prefer positive and negative tests.

Examples:

- expected success path
- missing file
- invalid metadata
- disconnected tenant
- changed package version
- stale selection
- retry path
- cache hit path
- cache miss path

A fix is not complete just because the happy path works.

---

## Working With Existing User Changes

If the working tree contains unrelated user changes:

- do not discard them
- do not overwrite them
- do not stage them accidentally
- do not reset them without explicit permission

If a conflict exists between requested work and local modifications, stop and explain the conflict before proceeding.

---

## Context Files

Before significant work, read:

`AGENTS.md`

`PROJECT_STATUS.md`

Use `PROJECT_STATUS.md` for:

- current version
- completed work
- recent architectural decisions
- deferred work
- release plan
- roadmap

If repository state contradicts `PROJECT_STATUS.md`, trust the actual repository state and flag the documentation as stale.

---

## Communication Style

When reporting work:

- be concise
- state what changed
- state which files changed
- state what was tested
- state any remaining risk
- show the relevant diff summary when useful

Do not claim something is fixed unless it was actually validated.