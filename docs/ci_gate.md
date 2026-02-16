# CI Gate Runbook (v0.3.0)

This document defines the required automated test checks for the v0.3.0 release line and how to enforce them.

## Required Checks

The following GitHub checks are required on pull requests to protected branches:

- `Core (py3.10)`
- `Core (py3.11)`
- `Core (py3.12)`
- `Core (py3.13)`
- `Core (py3.14)`
- `Optional deps (py3.14)`

These names must remain stable because branch protection rules match check names literally.

## Branch Protection / Ruleset Setup

Apply to both `development` and `main`:

1. Require a pull request before merging.
2. Require status checks to pass before merging.
3. Add all checks listed in **Required Checks**.
4. Enable "Require branches to be up to date before merging" (recommended).
5. Save and verify policy by opening a test PR.

## Workflow and Environment Mapping

- GitHub workflow file: `.github/workflows/tests.yml`
- Core matrix environments:
  - `py310`, `py311`, `py312`, `py313`, `py314`
- Optional environment:
  - `py314-optional`
- Tox configuration source:
  - `tox.ini`

## Failure Triage Policy (v0.3.0)

- Any failing required check blocks merge.
- Core failures are always blocking.
- Optional dependency failures in `Optional deps (py3.14)` are blocking for v0.3.0.

## Optional Dependency Scope Policy

For v0.3.0, optional dependency coverage is intentionally scoped to a single lane:

- `py314-optional`

Do not add a full optional matrix during v0.3.0 unless policy is explicitly updated.

## First Green Baseline Record

Fill this section when the first full required-check run is green:

- Date (UTC): `TBD`
- Commit SHA: `TBD`
- Workflow run URL: `TBD`
- Trigger type: `TBD` (`pull_request`, `push`, or `workflow_dispatch`)

## Local Commands

```bash
# Core suite (current interpreter)
QT_QPA_PLATFORM=offscreen pytest -q

# Multi-version gate (local, if interpreters are installed)
tox -e py310,py311,py312,py313,py314

# Optional dependency lane
tox -e py314-optional
```
