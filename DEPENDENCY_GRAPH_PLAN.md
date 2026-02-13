# Dependency Graph + Smart Cache Implementation Plan

## Design Decisions

### Core Design
- **SmartCache.INVALID = object()** - Sentinel to distinguish "not cached" from "cached None"
- **Deepcopy timing**: Lazy (on retrieval) - deepcopy when `C()` returns
- **Dynamic dependencies**: Re-track on every eval (correctness first, optimize later)

### Exception Hierarchy
```
RuntimeError
└── PyCellSheetError (base)
    └── CircularRefError
```

### Settings & UX
1. **Recalc mode**: Auto/manual (user option)
2. **Dirty visualization**: Refresh icon (circular arrows)
3. **Batch recalc**: Yes - "Recalculate All" button (Ctrl+Shift+F9)
4. **Persistence**: Rebuild dep graph on load (safer, simpler)
5. **Performance target**: 100 cells/sheet × 10 sheets = 1000 cells

---

## Architecture

### 1. DependencyGraph
- **Forward edges**: `dependencies[A2] = {A1}` (A2 depends on A1)
- **Reverse edges**: `dependents[A1] = {A2}` (A1 is depended on by A2)
- **Dirty flags**: `dirty = {A1, A2, A3}` (cells needing recalc)
- **Cycle detection**: DFS with recursion stack tracking

### 2. SmartCache
- **INVALID sentinel**: Distinguish "not cached" from "cached None"
- **Dependency-aware**: Check if dependencies are dirty before returning cache hit
- **Lazy deepcopy**: Store original in cache, deepcopy on retrieval

### 3. DependencyTracker (Context Manager)
- **Thread-local storage**: Track currently evaluating cell
- **Context manager**: `with DependencyTracker.track(key):`
- **Automatic recording**: When `C("A1")` is called, record dependency

---

## Implementation Phases

### Phase 1: Foundation (Week 1-2) ← **CURRENT**
**Order: Tests → Implementation → Integration**

#### 3. Write Tests First (TDD)
- `pyspread/lib/test/test_dependency_graph.py`
  - Test add/remove dependencies
  - Test cycle detection (simple, complex, self-reference)
  - Test dirty flag propagation
  - Test get_all_dependencies (transitive closure)

- `pyspread/lib/test/test_smart_cache.py`
  - Test INVALID sentinel (vs None)
  - Test cache hit/miss
  - Test dependency-based invalidation
  - Test clear()

- `pyspread/lib/test/test_exceptions.py`
  - Test CircularRefError formatting
  - Test cycle path display

#### 1. Implement Phase 1 Files
- `pyspread/lib/exceptions.py` - Exception hierarchy
- `pyspread/lib/dependency_graph.py` - DependencyGraph class
- `pyspread/lib/smart_cache.py` - SmartCache class

#### 2. Integrate into CodeArray
- Modify `CodeArray.__init__` - Add dep_graph and cache
- Modify `CodeArray.__setitem__` - Invalidate dependents, not all caches
- Modify `CodeArray.__getitem__` - Use cache, track dependencies
- Modify `C()`, `R()`, `Sh()` - Call `dep_graph.add_dependency()`

### Phase 2: Reference Parser Integration (Week 2-3)
- Wrap evaluation in `DependencyTracker.track()`
- Test cross-sheet dependencies
- Test range dependencies

### Phase 3: UI Integration (Week 3-4)
- Add recalc_mode setting (auto/manual)
- Add F9 / Ctrl+Shift+F9 shortcuts
- Add dirty cell visualization (refresh icon)
- Add "Auto Recalculate" toggle in menu

### Phase 4: Advanced Features (Week 4-5)
- Circular dependency error messages with path
- Topological sort for batch recalculation
- Manual mode: show "#DIRTY" in cells

### Phase 5: Optimization (Week 5-6)
- Static analysis for simple refs (skip dynamic tracking)
- Benchmark vs current approach
- Profile and optimize hot paths

---

## Example: Current vs Future

### Current (No Dep Graph)
```python
# User edits B1
CodeArray.__setitem__((1, 0, 0), "42")
    → self.result_cache = {}  # ❌ CLEARS ALL CACHES

# User accesses A4
CodeArray.__getitem__((0, 3, 0))
    → Not in cache, eval A4
    → A4 calls C("A3") → eval A3
    → A3 calls C("A2") → eval A2
    → A2 calls C("A1") → eval A1
    # ❌ Every access re-evaluates entire chain!
```

### Future (With Dep Graph + Smart Cache)
```python
# User edits B1
CodeArray.__setitem__((1, 0, 0), "42")
    → dep_graph.remove_cell(B1)
    → cache.invalidate(B1)
        → mark_dirty(B1)
        → Propagate to B2, B3 only
    # ✅ A1-A4 caches remain valid!

# User accesses A4
CodeArray.__getitem__((0, 3, 0))
    → cache.get(A4)
    → Check: A4 dirty? No
    → Check: dependencies dirty? No
    → ✅ CACHE HIT → return deepcopy(12)
    # ✅ No re-evaluation!
```

---

## Key Code Snippets

### CircularRefError Display
```python
# In cell:
"CircularRefError"

# In tooltip:
"Circular reference: A1 → A2 → A3 → A1 (cycle)"
```

### DirtyCell Display (Manual Mode)
```python
# In cell:
"#DIRTY"

# In tooltip:
"Cell needs recalculation. Press F9 to recalculate."
```

### Cache with INVALID Sentinel
```python
cached = self.cache.get(key)
if cached is not SmartCache.INVALID:
    return deepcopy(cached)  # Can be None (valid)
```

---

## Files to Create

### New Files
- `pyspread/lib/exceptions.py`
- `pyspread/lib/dependency_graph.py`
- `pyspread/lib/smart_cache.py`
- `pyspread/lib/test/test_dependency_graph.py`
- `pyspread/lib/test/test_smart_cache.py`
- `pyspread/lib/test/test_exceptions.py`

### Modified Files
- `pyspread/model/model.py` - CodeArray integration
- `pyspread/lib/pycellsheet.py` - Formatter, DependencyTracker, DirtyCell
- `pyspread/settings.py` - Add recalc_mode setting
- `pyspread/actions.py` - Add recalculate actions
- `pyspread/grid.py` - Add on_recalculate methods
- `pyspread/main_window.py` - Add recalc mode toggle

---

## Open Questions Resolved

1. ✅ **Recalc Strategy**: Make it an option (auto/manual)
2. ✅ **Dirty Visualization**: Refresh icon (circular arrows)
3. ✅ **Batch Recalc**: Yes - Ctrl+Shift+F9
4. ✅ **Persistence**: Rebuild on load
5. ✅ **Performance Target**: 1000 cells

---

## Next Steps

**Current Status**: About to start Phase 1 in order 3-1-2:
1. ✅ **Write tests** (TDD approach)
2. ⏳ **Implement files** (make tests pass)
3. ⏳ **Integrate** (CodeArray changes)

**Paused for**: User remembering what functionality to delete 💀
