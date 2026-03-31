# Part Search Tool

**Version:** 2.0.0 (refactor target)
**Platform:** Windows 10/11

---

## What This Tool Does

A desktop application for Oceaneering team members to:

1. **Search PSFT part and inventory data** via the Denodo REST API — without needing SQL knowledge or direct database access.
2. **Import a Bill of Materials (BOM) CSV** and see inventory levels across all relevant Business Units in a pivot view.
3. **Search indexed local and mapped-drive files** by filename (secondary feature).

The goal is to eliminate the time spent manually pulling, cross-referencing, and compiling part and inventory data from PeopleSoft. Any team member should be able to open this tool and get answers without training.

---

## Who Uses It

Non-technical users on the engineering and operations team. The UI must be operable by someone who has never seen it before, without reading a manual.

---

## Requirements

- Windows 10 or 11
- Python 3.11+
- Access to the Denodo REST API (internal network or VPN)
- Credentials stored in Windows Credential Manager (set up on first run)

### Python Dependencies

```
PyQt5==5.15.11
pandas==2.2.2
requests==2.32.3
keyring==25.2.1
openpyxl==3.1.2
urllib3==2.2.1
pywin32==306
```

Install with:
```bash
pip install -r requirements.txt
```

---

## First-Time Setup

1. Clone or download the repository.
2. Install dependencies: `pip install -r requirements.txt`
3. Run `python main.py`
4. On first launch, a credentials dialog will appear. Enter your Denodo username and password. These are stored securely in Windows Credential Manager — never in a file.

---

## Project Structure

```
part_search/
├── main.py                  # Entry point — 20 lines max
├── config.py                # All constants, defaults, paths
├── requirements.txt
├── data/
│   └── bus_list.json        # BU codes → descriptions (edit here to add/remove BUs)
├── core/
│   ├── denodo.py            # All Denodo API calls, filter builders, chunking logic
│   ├── file_index.py        # SQLite index builder for local file search
│   └── file_search.py       # File search worker thread and pane
├── ui/
│   ├── app_window.py        # Main window and tab navigation
│   ├── psft_pane.py         # PSFT search form and BU selector
│   ├── bom_pane.py          # BOM import and pivot results view
│   ├── results_widget.py    # Shared results table, export, copy
│   └── theme.py             # Dark/light/system theme application
├── utils/
│   ├── io_helpers.py        # JSON I/O, app paths, bootstrap
│   ├── auth.py              # Keyring credential management
│   └── normalize.py         # DataFrame column normalization helpers
└── tests/
    ├── test_smoke.py
    ├── test_denodo_filters.py
    └── test_normalize.py
```

---

## How to Add or Remove a Business Unit

Edit `data/bus_list.json`. Add or remove entries in the format:

```json
{
  "BUXXX": "Full BU Description Here"
}
```

No code change is required. The app loads this file at startup.

---

## How to Update the Denodo View

1. Update `config.UNIFIED_VIEW` with the new view name.
2. Update `config.DESIRED_HEADERS` if column names have changed.
3. Check `core/denodo.py` — the `normalize_keys()` call maps raw API column names to display names.
4. Run `python -m pytest tests/test_denodo_filters.py` to verify filter logic still works.

---

## What Is Intentionally Not In This Tool

| Feature | Reason Excluded |
|---|---|
| SQL query window | Developer debug tool; not useful to end users |
| ENOVIA / Online PLM integration | Requires IT permissions not currently granted |
| Serial barcode scanner | Not connected to any active workflow |
| Direct PeopleSoft access | Routed through Denodo by design |

---

## Known Limitations

- Inventory data accuracy is approximately ±24 hours (Denodo cache).
- File search covers only indexed local drives and mapped network drives. PDM vault contents require SolidWorks PDM Search directly.
- The following PSFT features are planned but not yet implemented: Where-Used search, BOM/Item-Contains review, Last Purchased/Ordered, NC/Deviation search, Rental Fleet search.

---

## Support and Maintenance

See `CONTRIBUTING.md` for how to make changes safely.
See `AGENTS.md` for how to coordinate AI-assisted development sessions.
See `TASKS.md` for the current work queue and status.