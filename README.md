# PyCellSheet

PyCellSheet is a fork of pyspread 2.3.1 that aims for a middle ground between:
- conventional spreadsheet behavior, and
- pyspread's fully Pythonic model.

The core design choice is copy-priority semantics: cell references return deep-copied values by default, so cells behave more like independent spreadsheet values.

## Project status

- Base fork point: `01500b4` (pyspread 2.3.1)
- v0.0.5 PoC tag point: `a7eda84`
- Current release-candidate line: v0.1.0 (`HEAD`)

## Key v0.1.0 RC changes

- Dependency-aware recalculation (`DependencyGraph` + `SmartCache`)
- Circular reference detection with explicit error handling
- Dirty-cell state visualization and recalculation tooling
- Named sheets and sheet rename support
- Expanded spreadsheet function coverage across core modules
- Removal of frozen-cell and evaluation-timeout features

## Docs

- Release notes: `changelog.txt`
- Design note: `Application Design Note.txt`
- Developer guide: `CLAUDE.md`
- User manual: `pyspread/share/doc/manual/`

## Run

```bash
pip install -r requirements.txt
python -m pyspread
```
