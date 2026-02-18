# PyCellSheet

PyCellSheet is a fork of pyspread 2.3.1 that aims for a middle ground between:
- conventional spreadsheet behavior, and
- pyspread's fully Pythonic model.

The core design choice is copy-priority semantics: cell references return deep-copied values by default, so cells behave more like independent spreadsheet values.

## Project status

- Base fork point: `01500b4` (pyspread 2.3.1)
- v0.3.0: released
- Current development phase: v0.4.0 internal-semantics cleanup

## Key changes

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

- `v0.4.0`: internal semantics cleanup (naming/refactor, warning contracts, reproducibility policy, button-cell cache/dependency contract).
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
