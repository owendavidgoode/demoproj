# CONTRIBUTING.md — Part Search Tool

## Purpose of This File

This document tells any contributor — human or AI — how to work on this codebase safely. Read it completely before making any changes. If something is unclear, open a task in `TASKS.md` to clarify it rather than guessing.

---

## Architecture Rules (Non-Negotiable)

These rules exist so that any module can be read, tested, or replaced without understanding the entire codebase.

| Layer | Folder | Rule |
|---|---|---|
| Business logic | `core/` | Zero PyQt5 imports. Testable without a display. |
| User interface | `ui/` | Calls `core/` for data. Never calls `requests` directly. |
| Utilities | `utils/` | Stateless helpers only. No imports from `core/` or `ui/`. |
| Configuration | `config.py` | Imported by everyone. Imports nothing from within the project. |

**If you are tempted to put logic directly in a UI event handler (`button.clicked.connect(lambda: ...)`), stop.** Put the logic in `core/`, write a test for it, then call it from the handler.

---

## The Denodo API

- **Base URL:** See `config.BASE_URL`
- **Auth:** HTTP Basic. Credentials stored in Windows Credential Manager under service name `"Denodo"`. Managed by `utils/auth.py`.
- **Single entry point:** All API calls go through `core/denodo.denodo_fetch_all_safe()`. Do not write raw `requests.get()` calls anywhere else.
- **Response format:** JSON, normalized to `pd.DataFrame` by the fetch function.
- **URL length limit:** 7,000 characters. The fetch function auto-chunks large `IN()` lists. Do not bypass this.
- **SSL:** Verification is disabled (`verify=False`) due to ZScaler MITM. This is intentional and known. Do not add a config toggle without a security review from IT.

---

## Key Functions Reference

| Function | Location | What it does |
|---|---|---|
| `items_lookup()` | `core/denodo.py` | Main PSFT search. Takes filter params, returns `(results_df, per_loc_df, raw_inv_df)`. |
| `build_itemmaster_filter()` | `core/denodo.py` | Converts UI inputs to a Denodo `$filter` string. |
| `denodo_fetch_all_safe()` | `core/denodo.py` | Robust GET with retry, auto-chunking, deduplication. |
| `normalize_keys()` | `utils/normalize.py` | Lowercases and snake_cases all DataFrame column names. |
| `normalize_bu()` | `utils/normalize.py` | Ensures BU column is present and clean. |
| `app_paths()` | `utils/io_helpers.py` | Returns dict of all file paths the app reads/writes. |
| `bootstrap_files()` | `utils/io_helpers.py` | Creates missing JSON config files on first run. |
| `get_basic_auth_header()` | `utils/auth.py` | Pulls credentials from keyring, returns auth header dict. |

---

## Data Files (Non-Code Configuration)

These files can be edited without touching Python. Prefer editing these over hardcoding values in source.

| File | Purpose | Who can edit |
|---|---|---|
| `data/bus_list.json` | BU code → description mapping | Anyone with a text editor |
| `%APPDATA%\PartSearch\user_prefs.json` | Per-user preferences | Written by the app; do not edit manually |
| `%APPDATA%\PartSearch\search_locations.json` | File search scope | Written by the app |
| `%APPDATA%\PartSearch\index_roots.json` | File index roots | Written by the app |

---

## UI Standards

Every UI change must follow these rules. There are no exceptions:

- **Plain English labels only.** No internal field names visible to users (e.g., `item_id` → "Part Number", `business_unit` → "Business Unit").
- **Enter key triggers search.** Every text input that can initiate a search must call `display_data` on `returnPressed`.
- **Status feedback is mandatory.** Every search must show a "Searching…" state while the worker thread runs and a result count or error when it finishes. A frozen window with no feedback is a bug.
- **Results table must support:** sort by any column, Ctrl+C to copy selection, Export to Excel button.
- **Tooltips on all non-obvious inputs.** One sentence max. Example: *"Enter one or more part numbers separated by semicolons."*
- **First-run credential prompt.** If no keyring credentials exist on startup, show the credentials dialog immediately with an explanation of what it's for.

---

## What Is Intentionally Missing

Do not re-add these without a documented decision to do so.

| Feature | Why It Was Removed |
|---|---|
| `SqlWindow` / `SQLHighlighter` | UI-facing SQL window is not useful to end users. `build_equivalent_sql()` is preserved in `core/denodo.py` for developer debugging only. |
| `items_lookup_old()` | Replaced by `items_lookup()`. All call sites updated. |
| `LocationPickerDialog` | Replaced by `SearchLocationsDialog`. Duplicate removed. |
| ENOVIA / PLM integration | Requires IT permissions not currently granted. All related code in `enovia_*.py` removed. |
| `SerialScanListener` | Not connected to any active user workflow. |

---

## Do Not Touch (Without Explicit Approval)

- **`utils/auth.py`** — Credential storage uses Windows Credential Manager. Never cache passwords in files, environment variables, or logs.
- **`denodo_fetch_all_safe()`** — The chunking and retry logic is carefully tuned for Denodo's URL length limits and ZScaler behavior. Changes require tests.
- **`verify=False` in the requests session** — Required. Do not change.
- **`data/bus_list.json`** — Updating entries is always fine. Changing the file format (structure, keys) requires updating `config.load_bus_list()` and testing.

---

## Testing

```bash
pip install -r requirements.txt
python -m pytest tests/ -v
```

- Tests require **no network access** and **no display** (headless Qt).
- `QApplication` in tests uses `sys.argv = ['']`.
- Every new function in `core/` or `utils/` must have at least one test.
- UI tests are smoke tests only (does it open without crashing).

### Definition of a Passing Test Suite
All tests pass. No new warnings introduced. Smoke test (`tests/test_smoke.py`) passes.

---

## Branch and Commit Conventions

### Branch Names
```
feat/short-description        # New feature
fix/short-description         # Bug fix
refactor/short-description    # Structural change, no behavior change
docs/short-description        # Documentation only
```

### Commit Messages
```
refactor: remove items_lookup_old and update BOM import call site
feat: add first-run credentials prompt on startup
fix: normalize_keys handles missing inventory_item_id column
docs: update TASKS.md block 1 status
```

---

## Definition of Done

A task is complete when:

1. The smoke test passes (`tests/test_smoke.py`).
2. The feature-specific test passes (or a new test was written and passes).
3. No new commented-out code was added. Comments explain *why*, not *what*.
4. `TASKS.md` is updated to reflect the new status.
5. The PR or commit description explains *why the change was made*, not just what changed.

---

## Questions and Ambiguities

If something in the code or requirements is unclear:

1. Check `ARCHITECTURE.md` first.
2. If still unclear, add a comment in `TASKS.md` under the relevant task with the `[?]` prefix.
3. Do not make an assumption and proceed — flag it and wait for clarification.