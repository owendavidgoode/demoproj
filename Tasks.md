# TASKS.md — Part Search Tool Work Queue

## How to Use This File

- Update status as work progresses. Do not delete completed tasks — they are the project history.
- Every AI session should reference the specific task block being worked on.
- If a task is blocked, document the reason and the next action needed to unblock it.
- New tasks go at the bottom of the relevant block, or in a new block if they don't fit.

## Status Key

```
[ ]  Not started
[~]  In progress
[x]  Complete
[!]  Blocked — reason documented in task notes
[?]  Needs clarification before work can begin
```

---

## BLOCK 0 — Repository Setup

- [ ] 0.1 Initialize git repository
  - Agent role: N/A (human task)
  - Files affected: entire project folder
  - Done when: `git init` complete, baseline committed with message `"baseline: pre-refactor"`, remote repo created (GitHub recommended)
  - Notes: All subsequent tasks should be done on branches. Nothing goes directly to `main`.

- [ ] 0.2 Create `requirements.txt` with pinned versions
  - Agent role: REFACTOR AGENT
  - Files affected: `requirements.txt` (new)
  - Done when: File exists with pinned versions for PyQt5, pandas, requests, keyring, openpyxl, urllib3, pywin32
  - Notes: Run `pip freeze` in the working environment to capture actual versions.

- [ ] 0.3 Create `tests/` folder with smoke test
  - Agent role: TEST AGENT
  - Files affected: `tests/__init__.py`, `tests/test_smoke.py` (new)
  - Done when: `python -m pytest tests/test_smoke.py` passes — app instantiates without crashing
  - Notes: Use headless QApplication (`sys.argv = ['']`). No network calls.

---

## BLOCK 1 — Dead Code Removal

Work on the original monolith file before splitting. Each task is a separate commit.

- [ ] 1.1 Delete `items_lookup_old()`
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`
  - Done when: Function removed, all call sites verified to use `items_lookup()`, smoke test passes
  - Notes: BOM import path must be confirmed to use `items_lookup()` before deleting.

- [ ] 1.2 Delete `LocationPickerDialog`
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`
  - Done when: Class removed, all references point to `SearchLocationsDialog`, smoke test passes
  - Notes: `SearchLocationsDialog` is the current implementation. `LocationPickerDialog` is the duplicate to remove.

- [ ] 1.3 Remove all commented-out code blocks
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`
  - Done when: No block-commented code remains. Inline `# reason` comments are fine. Commented-out `"""..."""` code blocks are not.
  - Notes: Specific targets — lines 820–826 (old `set_last_indexed_now_for_root`), `"""vb_tools.addStretch(1)"""`, `#Added by Ajay` inline code blocks. Git preserves history.

- [ ] 1.4 Audit and remove unused imports
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`
  - Done when: All imports at the top of the file are verifiably used by remaining code. Confirmed unused: `random`, and potentially `copy`, `glob`, `fnmatch`, `hashlib`, `uuid` (audit required after 1.1–1.3).
  - Notes: Run after 1.1–1.3 so removed code doesn't create false "still used" signals.

- [ ] 1.5 Remove `SerialScanListener`
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`
  - Done when: Class and all references removed. Smoke test passes.
  - Notes: Nothing in the current UI connects to this class. If barcode scanning becomes a requirement, it will be re-added with a proper UI connection and tests.

- [ ] 1.6 Remove `SqlWindow`, `SQLHighlighter`, and "Show SQL" button from UI
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`
  - Done when: Classes removed from UI. `show_sql()` method removed from `DenodoQuery`. `build_equivalent_sql()` function preserved (not deleted — it has debug value). "Show SQL" button removed.
  - Notes: `build_equivalent_sql()` moves to `core/denodo.py` during BLOCK 3. It is a developer tool only.

- [ ] 1.7 Remove ENOVIA / PLM dead code
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`
  - Done when: `enovia_config()`, `ensure_enovia_db()`, `enovia_ingest_csv()`, `enovia_db_path()`, `DEFAULT_ENOVIA_POLICY`, `DEFAULT_ENOVIA_CACHE`, and all ENOVIA-related JSON bootstrap calls removed.
  - Notes: This feature requires IT permissions that are not currently available. Removal documented in `CONTRIBUTING.md` under "What Is Intentionally Missing."

---

## BLOCK 2 — Extract Data Out of Code

- [ ] 2.1 Move `POSSIBLE_BUS` dict to `data/bus_list.json`
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`, `data/bus_list.json` (new), `config.py` (new)
  - Done when: `bus_list.json` contains all BU entries, `config.load_bus_list()` loads it at startup, BU checkboxes render correctly, no `POSSIBLE_BUS` dict in source code.
  - Notes: Any team member should be able to add a BU by editing the JSON file without opening a Python file.

- [ ] 2.2 Move all constants to `config.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `That_Search_Tool_3.py`, `config.py` (new)
  - Done when: `BASE_URL`, `UNIFIED_VIEW` (placeholder), `DESIRED_HEADERS`, `DEFAULT_BUS`, `APP_VERSION`, `DEFAULT_PREFS`, `DEFAULT_LOCATIONS`, `DEFAULT_INDEX_ROOTS`, `PN_REGEXES` all live in `config.py` and are imported from there.

---

## BLOCK 3 — Split Into Modules

Execute in order. After each task, run smoke test. If it passes, commit.

- [ ] 3.1 Create project package structure (empty files with `__init__.py`)
  - Agent role: REFACTOR AGENT
  - Files affected: Create folder structure and empty `__init__.py` files as defined in `ARCHITECTURE.md`
  - Done when: All folders and init files exist. `main.py` created with placeholder content.

- [ ] 3.2 Extract `utils/io_helpers.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `utils/io_helpers.py` (new), `That_Search_Tool_3.py`
  - Done when: `app_paths()`, `bootstrap_files()`, `_atomic_write_json()`, `save_json_atomic()`, `load_json_or_default()`, `_load_json_or_default()`, `read_index_roots_json()`, `write_index_roots_json()`, `read_search_locations()`, `write_search_locations()`, `maybe_migrate_from_programdata()`, `local_now_short()` extracted. All callers updated to import from `utils.io_helpers`.

- [ ] 3.3 Extract `utils/normalize.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `utils/normalize.py` (new), `That_Search_Tool_3.py`
  - Done when: `normalize_keys()`, `normalize_bu()`, `normalize_part()`, `normalize_all_columns()`, `normalize_status_key()` extracted. All callers updated.

- [ ] 3.4 Extract `utils/auth.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `utils/auth.py` (new), `That_Search_Tool_3.py`
  - Done when: `get_basic_auth_header()` and the credential dialog extracted. All callers updated.

- [ ] 3.5 Extract `core/denodo.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `core/denodo.py` (new), `That_Search_Tool_3.py`
  - Done when: All Denodo fetch functions, filter builders, `items_lookup()`, `aggregate_inventory()`, `build_equivalent_sql()` extracted. No PyQt5 imports in this file.

- [ ] 3.6 Extract `ui/results_widget.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `ui/results_widget.py` (new), `That_Search_Tool_3.py`
  - Done when: `PandasModel`, sort proxy, `ResultsWidget` (table + export + copy) extracted.

- [ ] 3.7 Extract `ui/theme.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `ui/theme.py` (new), `That_Search_Tool_3.py`
  - Done when: `apply_theme()`, `_apply_dark_palette()`, `_apply_light_palette()`, `_win_apps_uses_light()` extracted.

- [ ] 3.8 Extract `ui/psft_pane.py` and `ui/bom_pane.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `ui/psft_pane.py` (new), `ui/bom_pane.py` (new), `That_Search_Tool_3.py`
  - Done when: PSFT search form and BOM import view extracted into separate pane files.

- [ ] 3.9 Extract `core/file_index.py` and `core/file_search.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `core/file_index.py` (new), `core/file_search.py` (new), `That_Search_Tool_3.py`
  - Done when: Index DB helpers, `IndexBuilderWorker`, `FileSearchWorker`, `FileSearchPane` extracted.

- [ ] 3.10 Create `ui/app_window.py` and finalize `main.py`
  - Agent role: REFACTOR AGENT
  - Files affected: `ui/app_window.py` (new), `main.py`
  - Done when: `DenodoQuery` renamed/refactored to `AppWindow`, tab navigation wired, `main.py` is ~20 lines, smoke test passes.
  - Notes: This is the final step of the split. After this task, `That_Search_Tool_3.py` should be empty or deleted.

---

## BLOCK 4 — Unified Denodo View

- [!] 4.1 Obtain unified view schema from Denodo team
  - Agent role: N/A (human task)
  - Blocked on: Response from Denodo admin / Ajay with the unified view name and exact column list
  - Next action: Email Denodo team requesting: (1) view name, (2) full column list with data types, (3) confirmation that it replaces the three-view join currently done in Python
  - Done when: `config.py` has `UNIFIED_VIEW = "actual_view_name"` and a confirmed column mapping

- [ ] 4.2 Rewrite `items_lookup()` to use unified view
  - Agent role: DENODO AGENT
  - Files affected: `core/denodo.py`, `config.py`
  - Inputs required: Task 4.1 must be complete. Unified view name and column schema required.
  - Done when: `items_lookup()` makes a single `denodo_fetch_all_safe()` call against `UNIFIED_VIEW`. `_fetch_inventory_chunked()`, `aggregate_inventory()`, and the separate controls fetch block removed. BOM import path verified end-to-end.
  - Notes: Function signature must not change. Return type `(results_df, per_loc_df, raw_inv_df)` must be preserved.

- [ ] 4.3 Update filter tests for unified view
  - Agent role: TEST AGENT
  - Files affected: `tests/test_denodo_filters.py`
  - Done when: Tests verify filter string generation against the unified view's column names.

---

## BLOCK 5 — UI Simplification

- [ ] 5.1 Rename all fields to plain English
  - Agent role: UI AGENT
  - Files affected: `ui/psft_pane.py`
  - Done when: "AND wildcards" → "Description contains (all)"; "Either/Or wildcards" → "Description contains (any of)"; "Part No." → "Part Number"; all `QLabel` text reviewed for jargon.

- [ ] 5.2 Replace BU checkbox grid with searchable list
  - Agent role: UI AGENT
  - Files affected: `ui/psft_pane.py`
  - Done when: `QListWidget` with checkboxes replaces the flat checkbox grid. Search/filter box narrows the list. "Select All / None" and "Save as My Defaults" buttons present.
  - Notes: Underlying `_get_checked_bus()` logic must return the same list of BU codes as before.

- [ ] 5.3 Add status bar
  - Agent role: UI AGENT
  - Files affected: `ui/app_window.py`
  - Done when: Status bar shows "Ready" on launch, "Searching…" during fetch, "{N} results" on completion, error message on failure.

- [ ] 5.4 Add first-run credentials prompt
  - Agent role: UI AGENT
  - Files affected: `ui/app_window.py`, `utils/auth.py`
  - Done when: On startup, if no keyring credentials exist, a dialog appears explaining what credentials are needed and why. User cannot search until credentials are set.

- [ ] 5.5 Add inline help tooltips to all inputs
  - Agent role: UI AGENT
  - Files affected: `ui/psft_pane.py`
  - Done when: Every input field has a `QToolTip` or `?` label with a one-sentence description of what to enter.

---

## BLOCK 6 — PDM Search Integration

- [ ] 6.1 Research SolidWorks PDM Search CLI / COM interface
  - Agent role: N/A (human task)
  - Done when: One of the following confirmed — (A) PDM Search executable accepts CLI arguments, path documented; (B) PDM COM API (`ConisioLib.EdmVault`) accessible and can run queries; (C) Neither is viable, documented in `CONTRIBUTING.md` under limitations.
  - Notes: Check `%programfiles%\SolidWorks PDM\` for executables. Try `win32com.client.Dispatch("ConisioLib.EdmVault")` in a test script.

- [ ] 6.2 Implement PDM Search integration (if 6.1 confirms viability)
  - Agent role: DENODO AGENT (adapted for PDM)
  - Files affected: `core/pdm_search.py` (new), `ui/file_search_pane.py`
  - Inputs required: Task 6.1 must confirm the access method (CLI or COM)
  - Done when: "Search PDM" action available in File Search tab, returns results as a DataFrame, displayed in `ResultsWidget`.

- [ ] 6.3 Add permanent limitation notice to File Search tab
  - Agent role: UI AGENT
  - Files affected: `core/file_search.py` (FileSearchPane)
  - Done when: Persistent label visible in File Search pane: "Searches indexed local drives and mapped network drives only. For PDM vault contents, use SolidWorks PDM Search directly."
  - Notes: Do this regardless of 6.2 outcome. Honest labeling of tool capabilities.

---

## BLOCK 7 — Documentation and Handoff

- [ ] 7.1 Add docstrings to all public functions in `core/` and `utils/`
  - Agent role: REFACTOR AGENT
  - Files affected: All files in `core/` and `utils/`
  - Done when: Every public function has a one-line summary, Args, Returns, and Raises section.

- [ ] 7.2 Write `ARCHITECTURE.md` — verify against final module structure
  - Agent role: TASK AGENT / human review
  - Done when: `ARCHITECTURE.md` accurately reflects the final file structure after BLOCK 3 is complete. Data flow diagrams updated.

- [ ] 7.3 Final README review
  - Agent role: N/A (human task)
  - Done when: `README.md` accurately describes setup, usage, and limitations. Verified by someone who was not involved in writing the code.

---

## FUTURE / BACKLOG

These are documented requirements not yet scheduled.

- [ ] F.1 Where-Used part search
- [ ] F.2 Item-Contains (BOM Review) part search
- [ ] F.3 Last Purchased / Ordered part search
- [ ] F.4 NC / Deviation part search
- [ ] F.5 Rental fleet part search
- [ ] F.6 Barcode / serial scanner input (reconnect SerialScanListener with proper UI integration)
- [ ] F.7 Currency conversion toggle in BOM pivot view

---

## COMPLETED

*Tasks move here when done. Do not delete — this is project history.*

*(none yet)*