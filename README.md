# PyCellSheet

PyCellSheet is a fork of pyspread 2.3.1 that aims for a middle ground between:
- conventional spreadsheet behavior, and
- pyspread's fully Pythonic model.

The core design choice is copy-priority semantics: cell references return deep-copied values by default, so cells behave more like independent spreadsheet values.

## Project status

- Base fork point: `01500b4` (pyspread 2.3.1)
- v0.4.0: released
- Current development phase: v0.5.0 parser/spill feature expansion

## Key changes

### v0.5.0 (Parser + Spill Feature Expansion)

- Landed parser selector/migrator and spill-conflict feature line foundations.
- Tightened `.pycs` persistence contracts to named `[sheet_scripts]` headers only and canonical parser settings.
- Removed compatibility fallbacks that no longer match v0.5 behavior (numeric sheet-script headers and legacy parser-settings keys).
- Removed legacy `execute_macros()` alias surface; `execute_sheet_script()` is the canonical model entrypoint.
- Fixed `COUNTUNIQUE` behavior for ranges containing `EmptyCell` (no unhashable sentinel failure).
- Fixed `DataArray.data` setter contract to support direct dict assignment and preserve setter call compatibility.
- Fixed `PycsWriter.__len__` to match emitted save-stream length for accurate progress metadata.
- Migrated bundled `share/examples/*.pysu` files to v0.5 loader contracts (`[sheet_scripts]` + `[parser_settings]`).
- Sheet-name rule is now explicitly documented and enforced: no empty, whitespace-only, or control-character names.

### v0.4.0 (Internal Semantics Cleanup)

- Continued internal semantics cleanup and docs/test alignment.
- Added high-impact test coverage for command undo/redo paths, `.pycs` parser/writer edge cases, and model metadata/parser contracts.
- Hardened Sheet Script stream handling to always restore process stdio on interruption/failure paths.
- Exposed warning-marker details directly in cell tooltips for faster diagnosis.
- API docs pass cleaned normal-build warnings; strict cross-reference cleanup remains a later docs-quality task.

### v0.3.0 (Release Engineering)

- Release-engineering/safety baseline completion for the release line.
- Multi-version CI gate and tox matrix formalized as required checks.
- Sheet Script draft lifecycle safeguards stabilized across save/new/open/quit transitions.
- Baseline contract coverage expanded for parser modes, migration/alias compatibility, workflows, and runtime helper behavior.

### v0.2.0 (Stabilization)

- Recalc UX baseline improvements (clearer status feedback and dirty-indicator consistency).
- Sheet Script workflow expansion (apply-all behavior, trusted/untrusted load handling, safe-mode approval consistency).
- Dependency/recalculation stabilization (dynamic-reference correctness and transitive traversal hardening).
- Save/load robustness for sheet names and per-sheet script headers, plus broader stabilization test coverage.

### v0.1.0 (Core Feature Release)

- Dependency-aware recalculation (`DependencyGraph` + `SmartCache`)
- Circular reference detection with explicit error handling
- Dirty-cell state visualization and recalculation tooling
- Named sheets and sheet rename support
- Expanded spreadsheet function coverage across core modules
- Removal of frozen-cell and evaluation-timeout features

### v0.0.5 (Proof-of-Concept)

- Fork from pyspread 2.3.1 with PyCellSheet direction established.
- Copy-priority core model introduced (`EmptyCell`, `PythonCode`, `Range`, `RangeOutput`).
- Initial expression parsing modes and reference parsing pipeline implemented.
- Per-sheet init scripts, `.pycs` workflow, and first spreadsheet helper/function layers added.

## Roadmap

- `v0.5.0`: parser and spill expansion (parser selector/migrator, RangeOutput conflict semantics, staged `pycel` integration).
- `v0.6.0`: architecture upgrades (data model rewrite, compile caching, dynamic row/column sizing, recalc/warning UX completion).

## Docs

- Release notes: `changelog.txt`
- Design note: `Application Design Note.txt`
- Developer guide: `CLAUDE.md`
- User manual: `pycellsheet/share/doc/manual/`

## Run

```bash
pip install -r requirements.txt
python -m pycellsheet
```

## Testing

CI gate policy and required-check contract are documented in `docs/ci_gate.md`.

```bash
# Core suite (current interpreter)
QT_QPA_PLATFORM=offscreen pytest -q

# Multi-version gate (local, if interpreters are installed)
tox -e py310,py311,py312,py313,py314

# Optional dependency coverage (latest line)
tox -e py314-optional
```

Current active-interpreter baseline: `941 passed`.
