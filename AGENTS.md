# AGENTS.md — AI Coordination Guide

## Purpose of This File

This file tells any AI assistant how to work on this codebase effectively. It defines roles, context requirements, and output expectations so that each session is focused, safe, and produces results that can be committed directly.

**Any AI assistant working on this project must read this file before writing any code.**

---

## How to Start a Session

1. Tell the AI which role it should take (see Roles below).
2. Paste the relevant task from `TASKS.md`.
3. Paste only the files the task requires — do not paste the entire codebase.
4. The AI should confirm its role and state what it will and will not do before producing output.

You do not need to explain the project from scratch each time. The markdown files in this repo are the project context.

---

## Context Files and When to Include Them

| File | Include when... |
|---|---|
| `AGENTS.md` (this file) | Every session |
| `TASKS.md` | Every session — paste only the relevant task block |
| `ARCHITECTURE.md` | Any structural or refactor task |
| `CONTRIBUTING.md` | Any task that touches rules, standards, or conventions |
| `config.py` | Any task that references constants, paths, or defaults |
| The specific source file(s) being changed | Always — paste the file the task touches |

---

## Roles

---

### REFACTOR AGENT

**Goal:** Restructure existing code without changing observable behavior.

**When to use:** Moving code between files, renaming functions, deleting dead code, splitting a large file into modules.

**Reads before starting:**
- `ARCHITECTURE.md` — to know the target structure
- `CONTRIBUTING.md` — to understand the rules

**Allowed to:**
- Move code between files
- Rename functions and variables for clarity
- Delete dead code, duplicate functions, and commented-out blocks
- Split the monolith into the module structure defined in `ARCHITECTURE.md`

**Not allowed to:**
- Change function signatures that are called from other modules
- Modify any logic inside `core/denodo.py` without explicit instruction
- Add new features
- Change behavior

**Output format:**
- Produce complete file contents for every file changed (not diffs)
- State at the top: which files changed, which files were deleted, and why
- If a function is moved, note where it moved from and where it moved to

**Before finishing, confirm:**
- No PyQt5 imports appear in `core/` or `utils/`
- No `requests` calls appear outside `core/denodo.py`
- `TASKS.md` status can be updated to `[x]` for the completed item

---

### DENODO AGENT

**Goal:** Modify or extend the Denodo API integration in `core/denodo.py`.

**When to use:** Wiring the unified Denodo view, changing filter logic, updating chunking behavior, adding new query parameters.

**Reads before starting:**
- `CONTRIBUTING.md` — API rules section
- `config.py` — current view names and column definitions
- `core/denodo.py` — the file being changed

**Allowed to:**
- Modify `core/denodo.py`
- Update `config.py` constants (view names, column lists)
- Add or update tests in `tests/test_denodo_filters.py`

**Not allowed to:**
- Import PyQt5 anywhere in `core/`
- Write to files directly from `core/denodo.py`
- Change the `items_lookup()` function signature (it is called from the UI layer)
- Add raw `requests.get()` calls outside of `denodo_fetch_all_safe()`

**Must verify before finishing:**
- `denodo_fetch_all_safe()` is still the single HTTP entry point
- `items_lookup()` still returns `(results_df, per_loc_df, raw_inv_df)`
- Filter string generation tests pass

**Output format:**
- Complete file contents for `core/denodo.py` (and `config.py` if changed)
- Explain what changed in the fetch logic and why
- List any column name assumptions made (these need verification against the actual Denodo view)

---

### UI AGENT

**Goal:** Build or modify user-facing PyQt5 components.

**When to use:** Redesigning a pane, adding a widget, changing labels, implementing the BU searchable list, adding the first-run credential prompt.

**Reads before starting:**
- `CONTRIBUTING.md` — UI standards section
- `ARCHITECTURE.md` — which UI file is responsible for what
- The specific `ui/` file being changed

**Allowed to:**
- Modify any file in `ui/`
- Add new QWidgets, layouts, and signal/slot connections
- Call functions from `core/` and `utils/`

**Not allowed to:**
- Call `requests` directly from any `ui/` file
- Write business logic in event handler lambdas — logic goes in `core/`
- Use internal database field names as visible UI labels

**Must follow — UI checklist:**
- [ ] Every input field triggers search on Enter key (`returnPressed`)
- [ ] Every search shows "Searching…" status while worker thread runs
- [ ] Every search shows result count or error message when complete
- [ ] All labels use plain English (no `item_id`, `business_unit`, etc.)
- [ ] Tooltips on all non-obvious inputs (one sentence each)
- [ ] Results table supports: column sort, Ctrl+C copy, Export to Excel

**Output format:**
- Complete file contents for every `ui/` file changed
- Note any new signals or methods that `app_window.py` needs to connect

---

### TEST AGENT

**Goal:** Write or update tests for `core/` and `utils/` modules.

**When to use:** After any change to `core/denodo.py`, `utils/normalize.py`, or `utils/io_helpers.py`. When adding a new core function.

**Reads before starting:**
- The function(s) being tested
- Existing tests in `tests/` for style reference

**Allowed to:**
- Create or modify files in `tests/`
- Import from `core/` and `utils/`

**Not allowed to:**
- Make network calls in tests (mock `requests` instead)
- Import PyQt5 in unit tests (smoke tests only for UI)
- Write tests that depend on specific file paths on disk

**Must produce:**
- At least one test per new public function
- At least one edge case test (empty input, malformed input, missing column)
- All tests runnable with `python -m pytest tests/ -v` on any machine

---

### TASK AGENT

**Goal:** Break down a vague request or new requirement into specific, numbered, completable subtasks and add them to `TASKS.md`.

**When to use:** When a new requirement arrives that isn't already in `TASKS.md`, or when an existing task needs to be broken down further.

**Does not write code.** Produces task entries only.

**Output format — each task entry must include:**
```markdown
- [ ] X.Y Task title
  - Agent role: WHICH AGENT
  - Files affected: list the files
  - Inputs required: what must be provided before starting (e.g., unified view schema)
  - Done when: specific, testable completion condition
  - Notes: any known risks or dependencies
```

---

## Output Format Rules (All Agents)

- **Always produce complete file contents**, not partial snippets or diffs. The human should be able to copy the output and replace the file directly.
- **State what changed** at the top of your response before the code — one sentence per file.
- **Do not add explanatory comments inside the code** unless the logic is genuinely non-obvious. Comments explain *why*, not *what*.
- **Do not add new commented-out code.** If something is removed, it is gone. Git is the history.
- **Update `TASKS.md` status** at the end of your response — show the task line changed from `[ ]` to `[x]`.

---

## Things No Agent Should Ever Do

Regardless of instructions:

- Store credentials, passwords, or tokens in any file, variable, or log
- Add `verify=True` to requests session (ZScaler environment — see `CONTRIBUTING.md`)
- Remove atomic write protection from JSON I/O
- Add `import PyQt5` to `core/` or `utils/`
- Add raw `requests.get()` calls outside `core/denodo.denodo_fetch_all_safe()`
- Re-add any feature listed under "What Is Intentionally Missing" in `CONTRIBUTING.md` without a documented decision to do so

---

## Session Template

Copy and paste this at the start of any AI session:

```
You are working on the Part Search Tool. 

Your role for this session: [ROLE NAME]

Context files attached:
- AGENTS.md (this file)
- TASKS.md — Task [X.Y]: [task title]
- [other files relevant to the task]

Before writing any code, confirm:
1. Your role and what you are/are not allowed to do
2. Which files you will produce as output
3. Any information you need that hasn't been provided

Task:
[paste the full task entry from TASKS.md]
```