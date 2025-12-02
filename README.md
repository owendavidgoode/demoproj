# Local PDM/PLM Inventory Tool

A local-only CLI utility for indexing SolidWorks PDM (via mapped drives) and Aras/Enovia-style PLM (via browser automation), cross-referencing remote files against your local filesystem, and running read-only PeopleSoft/Denodo queries.

> **Status:** Prototype. PDM indexing and JSON output work end-to-end. PLM scraping is still mocked and must be wired to your real PLM UI before this is production-usable.

## Features (Current vs Planned)

- **PDM Indexing (working)**  
  Scan configured roots (e.g., mapped PDM vault drives) and capture file metadata (name, local path, relative path, size, timestamps).
- **PLM Scraping (scaffolded, mocked)**  
  Selenium-based login + scan scaffolding; currently yields mock PLM items. You must add site-specific selectors and traversal logic for your PLM.
- **Presence Checking (basic)**  
  Cross-reference PLM items against PDM by relative path / filename; outputs `present_locally: true/false` and aggregate match/miss stats.
- **PeopleSoft Search (working)**  
  Run read-only SQL via Denodo/ODBC; blocks write operations.
- **Local Search (working)**  
  Search the generated inventory JSON by name or path.

## Runtime Requirements

You will need, on the machine where you run this:

- **OS**
  - Windows 10/11 (64-bit), with access to:
    - Mapped PDM vault drives (e.g., `Z:\Vault\...`)
    - Network paths / archive drives you care about.
- **Python**
  - Python 3.10+ installed and on `PATH`.
- **Browser + Driver for PLM**
  - Google Chrome or Microsoft Edge installed.
  - Matching `chromedriver`/`msedgedriver` either:
    - Placed in the local `bin/` directory in this repo (recommended), or
    - Available on `PATH` (less portable if installs are blocked).
- **ODBC / Denodo**
  - `pyodbc` Python package (installed via `requirements.txt`).
  - A configured ODBC DSN that points at Denodo/PeopleSoft (for example `DenodoODBC`), created via the Windows ODBC Data Sources tool.
  - Network access from your machine to the Denodo/PeopleSoft endpoint.
- **Network Access**
  - Ability to reach your PLM web server from the machine and browser you use for scraping.

## Quick Start (PDM-only)

1. `make install`
2. `cp config/settings.example.json config/settings.json`
3. Set `"pdm.roots"` in `config/settings.json` to your mapped vault paths (e.g., `Z:\\Vault`).
4. Run `python -m src.cli.main index --pdm-only --force` to build `inventory.json` locally.
5. Re-run with `--dry-run` any time you want to check changes without writing output.

## Installation

From the project root:

```bash
# 1) Create virtualenv and install Python deps
make install

# 2) Copy and configure settings
cp config/settings.example.json config/settings.json
```

## Configuration

Then edit `config/settings.json`:

```json
{
  "pdm": {
    "roots": ["Z:\\Vault\\Designs", "C:\\Local\\Archive"]
  },
  "plm": {
    "url": "https://plm.example.com/innovator",
    "username": "",
    "password": "",
    "headless": false,
    "save_cookies": false
  },
  "peoplesoft": {
    "connection_string": "DSN=DenodoODBC;UID=user",
    "query_timeout": 30
  },
  "output": {
    "path": "inventory.json"
  }
}
```

> **Security note:** leave `plm.username` and `plm.password` empty in config and let the tool prompt you at runtime. This avoids storing credentials on disk. The example file is intentionally blank.

### What Works Today vs What You Must Wire Up

- **Works today**
  - PDM indexing over large directory trees (streamed to JSON).
  - Inventory writing with summary stats (`total_pdm`, `total_plm`, `matched`, `missing_locally`).
  - Local search and PeopleSoft/Denodo queries.
- **You must still wire up before this is “real”**
  - PLM login selectors and traversal in `src/indexer/plm.py` (currently uses mock data).
  - Confirm chromedriver/Edge driver is present and matches your browser version.
  - Validate Denodo/PeopleSoft DSN and permissions for your account.

## Usage

### Index PDM/PLM Files

```bash
# Index both PDM and PLM
python -m src.cli.main index --force

# Index PDM only (no browser needed)
python -m src.cli.main index --pdm-only --force

# Index PLM only with cookie persistence
python -m src.cli.main index --plm-only --save-cookies --force

# Dry run (no output written)
python -m src.cli.main index --dry-run

# With filters
python -m src.cli.main index --ext .sldprt --date-from 2024-01-01 --path-prefix /Vault/Projects
```

### Search PeopleSoft

```bash
# Run query from file
python -m src.cli.main search-ps queries/parts.sql

# Prompt for credentials (recommended)
python -m src.cli.main search-ps queries/parts.sql --prompt-creds
```

### Search Local Index

```bash
python -m src.cli.main search-local "part123"
```

### CLI Flags

| Flag | Description |
|------|-------------|
| `--config PATH` | Path to settings.json |
| `--verbose, -v` | Enable debug logging |
| `--force, -f` | Overwrite existing output |
| `--resume, -r` | Resume from checkpoint |
| `--dry-run` | Simulate without writing |
| `--pdm-only` | Index only filesystem/PDM |
| `--plm-only` | Index only PLM web UI |
| `--save-cookies` | Persist PLM session cookies |
| `--ext` | Filter by file extension |
| `--path-prefix` | Filter by path prefix |
| `--date-from` | Filter from date (ISO format) |
| `--date-to` | Filter to date (ISO format) |
| `--prompt-creds` | Prompt for PeopleSoft credentials |

## Development

- `make dev` — show CLI help (runs inside the venv).
- `make test` — run pytest suite in `tests/`.
- `make lint` — install and run Ruff over `src/` and `tests/`.
- `make build` — create `dist/local-inventory-tool.zip` with source + config template.
- `make clean` — drop venv, build artifacts, and `__pycache__` directories.

## Output Format

The inventory JSON contains:

```json
{
  "items": [
    {
      "name": "part123.sldprt",
      "remote_path": "/Vault/Projects/A/part123.sldprt",
      "remote_id": "1001",
      "local_path": "Z:\\Vault\\Projects\\A\\part123.sldprt",
      "created_at": "2024-01-15T10:30:00",
      "modified_at": "2024-06-20T14:22:00",
      "present_locally": true,
      "source": "plm"
    }
  ],
  "summary": {
    "total_items": 1500,
    "stats": {
      "total_pdm": 1200,
      "total_plm": 300,
      "matched": 280,
      "missing_locally": 20
    },
    "status": "completed"
  }
}
```

## PLM Scraping Setup

The PLM indexer is scaffolded but intentionally conservative. To make it talk to your real PLM:

1. **Run interactively once**
   - Set `"headless": false` in `config/settings.json`.
   - Run `python -m src.cli.main index --plm-only --force`.
   - A browser window should open at `plm.url`.
2. **Implement login DOM interactions**
   - Open `src/indexer/plm.py` and update `PLMIndexer.login()`:
     - Uncomment and adapt the `WebDriverWait(...).until(...)` and `find_element(...).send_keys(...)` lines to match your login page (username field, password field, submit button).
     - If you have MFA, leave enough `self._random_sleep()` time to enter the code manually on first run.
3. **Implement scan logic**
   - Replace the `mock_files` block in `PLMIndexer.scan()` with:
     - Navigation to the main search/list view.
     - A loop over rows/cards representing files.
     - For each row, extract `name`, logical `remote_path` (e.g., `/Vault/Folder/Subfolder`), and any stable `remote_id`.
     - Optional: follow pagination (next-page button) until no further pages exist.
4. **Test small**
   - Run with `--plm-only --dry-run -v` and add debug logging until you see realistic PLM items being yielded.
5. **Enable cookie reuse (optional)**
   - If allowed by policy, add `--save-cookies` once so a long-lived session can be reused on subsequent runs without re-entering MFA.

### Troubleshooting PLM Scraping

- **Login fails**: Check `login_form`, `username_field`, `password_field`, `login_button` selectors
- **No items found**: Verify `main_grid` and `grid_rows` match your PLM's table structure
- **Missing data**: Update `item_name`, `item_id`, `item_created`, `item_modified` selectors
- **Pagination stops**: Check `next_page` selector and disabled-state detection

## Architecture

```
src/
├── cli/main.py          # CLI entrypoint and command handlers
├── indexer/
│   ├── pdm.py           # Filesystem scanner for PDM mapped drives
│   └── plm.py           # Selenium-based PLM web scraper
├── search/
│   ├── local.py         # Inventory search
│   └── peoplesoft.py    # Denodo/PeopleSoft SQL executor
├── storage/
│   ├── inventory.py     # Streaming JSON writer/reader
│   └── checkpoint.py    # Resume state management
└── utils/
    ├── config.py        # Configuration loader
    ├── logging.py       # Logging setup
    └── validation.py    # Input sanitization
```

## Considerations for Non-LLM Developers

This codebase was developed with AI assistance. Here are key points for traditional development:

### Code Patterns Used

- **Generator-based streaming**: `PDMIndexer.scan()` and `PLMIndexer.scan()` yield items one at a time to handle 20TB+ vaults without memory issues
- **Context managers**: `InventoryWriter` uses `__enter__`/`__exit__` for safe file handling
- **Defensive error handling**: Stale element exceptions, timeouts, and partial failures are caught and logged

### Extending the PLM Scraper

To add support for a new PLM system:

1. Copy and adapt the Selenium element-finding calls inside `PLMIndexer.login()` and `PLMIndexer.scan()` to your PLM’s DOM.
2. If the PLM uses iframes, add `driver.switch_to.frame(...)` calls before locating elements.
3. For AJAX-heavy UIs, increase waits (e.g., `WebDriverWait(..., timeout=20)`) and wait on specific row/table locators.
4. For non-standard auth (SAML, OAuth), extend `login()` to follow redirects and only save cookies once fully authenticated.

### Testing Locally

```bash
# Run PDM indexer on a test directory
python -m src.cli.main index --pdm-only --config config/test.json --dry-run -v

# Test PLM connection without full scan
python -c "from src.indexer.plm import PLMIndexer; p = PLMIndexer({'url': '...', 'headless': False}); p.login()"
```

### Common Modifications

| Task | File | Function |
|------|------|----------|
| Add new CLI flag | `src/cli/main.py` | `main()` argparse setup |
| Change output schema | `src/storage/inventory.py` | `InventoryWriter.add_item()` |
| Add new filter type | `src/cli/main.py` | `apply_filters()` closure |
| Support new PLM | `src/indexer/plm.py` | Edit `PLMIndexer.login()` / `scan()` |
| Add presence matching logic | `src/cli/main.py` | Lines 184-200 |

### Dependencies

Core dependencies are minimal and chosen for portability:

- `selenium`: Browser automation (bundled chromedriver preferred)
- `pyodbc`: Database connectivity for PeopleSoft/Denodo
- `tqdm`: Progress bars (optional, graceful fallback)

No Playwright (blocked by policy), no heavy frameworks.

### Security Notes

- Credentials are prompted at runtime, not stored in config
- Cookies only persist with explicit `--save-cookies` flag
- SQL queries are checked for write operations before execution
- Paths are validated to prevent traversal attacks

## License

Internal use only.
