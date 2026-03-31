# ARCHITECTURE.md — Part Search Tool

## Purpose of This File

This document describes the intended structure of the codebase after refactoring. It is the authoritative reference for where code belongs and why. When in doubt about where something goes, the answer is here.

---

## Guiding Principle

**Each layer knows about the layers below it. No layer knows about the layers above it.**

```
ui/          ← knows about core/ and utils/
core/        ← knows about utils/ only
utils/       ← knows about nothing inside the project
config.py    ← knows about nothing inside the project
```

This means you can test `core/` without launching a UI. You can swap the UI framework without touching business logic. You can read any single file and understand it without reading everything else.

---

## Full Directory Structure

```
part_search/
│
├── main.py                  # Entry point. ~20 lines. No logic.
├── config.py                # All constants, defaults, paths. No logic.
├── requirements.txt         # Pinned dependency versions.
│
├── data/
│   └── bus_list.json        # BU code → description map. Edit here to add/remove BUs.
│
├── core/                    # Business logic. Zero PyQt5 imports.
│   ├── __init__.py
│   ├── denodo.py            # Denodo API: fetch, filter builders, chunking, normalization.
│   ├── file_index.py        # SQLite index builder. Crawls local/mapped drives.
│   └── file_search.py       # File search worker thread. Queries the index.
│
├── ui/                      # PyQt5 presentation layer. No direct API calls.
│   ├── __init__.py
│   ├── app_window.py        # Main QWidget window. Tab navigation. Wires panes together.
│   ├── psft_pane.py         # PSFT search form: part no., description, BU selector, filters.
│   ├── bom_pane.py          # BOM CSV import and per-BU pivot results view.
│   ├── results_widget.py    # Shared: PandasModel, sort proxy, table view, export, copy.
│   └── theme.py             # Dark / light / system theme via QPalette.
│
├── utils/                   # Stateless helpers. No imports from core/ or ui/.
│   ├── __init__.py
│   ├── io_helpers.py        # JSON read/write (atomic), app_paths(), bootstrap_files().
│   ├── auth.py              # Keyring get/set, credential dialog, Basic auth header.
│   └── normalize.py         # normalize_keys(), normalize_bu(), normalize_part().
│
└── tests/
    ├── test_smoke.py         # App opens without crashing.
    ├── test_denodo_filters.py # Filter string generation correctness.
    └── test_normalize.py     # normalize_keys, normalize_bu, normalize_part.
```

---

## Module Descriptions

### `main.py`
Entry point only. Calls `bootstrap_files()`, creates `QApplication`, applies theme, instantiates `AppWindow`, calls `show()`. Must not contain any logic that isn't setup. Target: 20 lines.

### `config.py`
Single source of truth for every constant, default, and path the application uses. Imported by every other module. Must not import from anywhere within the project. Contains:
- `BASE_URL`, `UNIFIED_VIEW`, `DESIRED_HEADERS`
- `DEFAULT_BUS`, `APP_VERSION`
- `_default_prefs()`, `load_bus_list()`, `load_config()`

### `data/bus_list.json`
Plain JSON file. The only place Business Unit codes and descriptions live. Format: `{"BUXXX": "Full description"}`. No code change needed to add or remove a BU.

---

### `core/denodo.py`
All interaction with the Denodo REST API. This is the most critical module in the project.

**Responsibilities:**
- `denodo_fetch_all_safe()` — single entry point for all HTTP GET calls. Handles retries, SSL, auto-chunking for URL length limits, deduplication.
- `build_itemmaster_filter()` — converts structured UI input (part IDs, description terms, quality codes) into a Denodo `$filter` string.
- `build_inv_filters()` — builds the inventory-side filter for BU and item ID matching.
- `items_lookup()` — orchestrates a full PSFT search: calls the unified view, applies filters, normalizes and returns a `(results_df, per_loc_df, raw_inv_df)` tuple.
- `aggregate_inventory()`, `aggregate_inventory_per_loc()` — roll up raw inventory rows into per-item and per-location summaries.
- `build_equivalent_sql()` — developer utility. Produces a human-readable SQL representation of what the API call is doing. Not exposed in the UI.

**Must not:** Import PyQt5. Write to files. Know anything about the UI.

### `core/file_index.py`
Manages the per-root SQLite index databases used by file search.

**Responsibilities:**
- `ensure_quick_index_db()` — creates and migrates the schema for a root's index DB.
- `IndexBuilderWorker` (QThread) — crawls a root directory tree, inserts file records into SQLite, emits progress signals.
- `IndexBuilderPane` (QWidget) — UI for managing index roots, triggering rebuilds, showing status.

**Note:** `IndexBuilderPane` is a QWidget and imports PyQt5. It lives in `core/` because it is tightly coupled to the index DB logic, but it is the one exception to the "no PyQt5 in core/" rule. If this feels wrong, it can be split into `core/file_index_db.py` (pure logic) and `ui/index_pane.py` (widget).

### `core/file_search.py`
Searches the SQLite index built by `file_index.py`.

**Responsibilities:**
- `FileSearchWorker` (QThread) — queries indexed DBs, optionally falls back to Windows Search index or live crawl. Emits results in chunks for UI responsiveness.
- `FileSearchPane` (QWidget) — search form, results table, source selector (Quick Index / Windows Index / Crawl).

**Known limitation:** Covers only locally indexed drives and mapped network drives. PDM vault contents are not accessible without additional IT permissions or SolidWorks PDM COM API integration.

---

### `ui/app_window.py`
The top-level `QWidget`. Owns the tab bar and `QStackedWidget`. Wires together `PsftPane`, `BomPane`, and `FileSearchPane`. Manages theme preference and window geometry persistence. Must not contain search logic.

### `ui/psft_pane.py`
The PSFT search form. Contains:
- Text inputs: Part Number, MFG Part Number, Description Contains (All), Description Contains (Any Of), Quality Codes, Minimum Quantity.
- BU selector: searchable `QListWidget` with checkboxes. Check All / None. Save as Default.
- Action buttons: Search, Import CSV, Set Credentials, Save Defaults, Reset.
- Connects user actions to `core/denodo.items_lookup()` via a worker thread.
- Displays results via `ui/results_widget.ResultsWidget`.

### `ui/bom_pane.py`
Handles BOM CSV import and the pivot view. Parses the uploaded CSV, extracts part numbers, calls `core/denodo.items_lookup()` for the full list, and renders a per-BU availability matrix.

### `ui/results_widget.py`
Reusable results display used by both `psft_pane.py` and `bom_pane.py`. Contains:
- `PandasModel` — `QAbstractTableModel` backed by a `pd.DataFrame`.
- `QSortFilterProxyModel` subclass — column sort with type-aware comparison (numeric, date, string).
- `ResultsWidget` — the full table view with column resize, Ctrl+C copy, and Export to Excel button.

### `ui/theme.py`
Self-contained theme module. `apply_theme(app, mode)` where mode is `"dark"`, `"light"`, or `"system"`. Reads Windows registry to detect system preference on Windows. No dependencies on any other project module.

---

### `utils/io_helpers.py`
All file I/O that isn't credentials. Atomic JSON writes (write to `.tmp`, then `os.replace`). `app_paths()` returns a dict of every path the app uses, derived from `%APPDATA%`. `bootstrap_files()` creates any missing JSON files with defaults on first run.

### `utils/auth.py`
Everything credential-related. `get_basic_auth_header()` pulls from Windows Credential Manager and returns an HTTP header dict. `CredentialDialog` is the QDialog for entering username and password. Credentials are never written to disk in plaintext.

### `utils/normalize.py`
Pure data transformation functions. No I/O, no UI, no API calls. `normalize_keys()`, `normalize_bu()`, `normalize_part()`, `normalize_all_columns()`, `normalize_status_key()`. Each function takes a DataFrame (or string) and returns a cleaned version.

---

## Data Flow: PSFT Search

```
User fills form (ui/psft_pane.py)
    │
    ▼
PsftPane._on_search()
    │
    ├── reads form inputs
    ├── calls core/denodo.items_lookup(item_ids, mfg_parts, and_terms, ...)
    │       │
    │       ├── build_itemmaster_filter() → $filter string
    │       ├── denodo_fetch_all_safe(UNIFIED_VIEW, headers, params)
    │       │       │
    │       │       └── HTTP GET → Denodo REST API
    │       │               └── returns JSON
    │       │
    │       ├── normalize_keys(df)          [utils/normalize.py]
    │       ├── normalize_bu(df)            [utils/normalize.py]
    │       ├── apply_qcode_filter(df)      [core/denodo.py]
    │       ├── apply_min_qty_filter(df)    [core/denodo.py]
    │       └── apply_bu_filter(df)         [core/denodo.py]
    │
    ▼
ResultsWidget.load(df)          (ui/results_widget.py)
    │
    ├── PandasModel(df)
    └── table renders results
```

---

## Data Flow: BOM Import

```
User clicks "Import CSV" (ui/psft_pane.py or ui/bom_pane.py)
    │
    ▼
QFileDialog → user selects .xlsx or .csv
    │
    ├── parse file → extract part numbers → list[str]
    ├── calls core/denodo.items_lookup(item_ids=part_list, ...)
    │       └── (same flow as PSFT Search above)
    │
    ▼
BomPane.render_pivot(results_df)
    └── pivot by BU, render in ResultsWidget
```

---

## Constraints Summary

| Rule | Enforced By |
|---|---|
| `core/` has no PyQt5 | Code review + import linter |
| All API calls via `denodo_fetch_all_safe` | Code review |
| Credentials only via `utils/auth.py` | Code review |
| BU list in `data/bus_list.json` | `config.load_bus_list()` is the only place it's read |
| No plaintext passwords anywhere | `utils/auth.py` uses keyring exclusively |
| Atomic JSON writes only | `utils/io_helpers._atomic_write_json()` |