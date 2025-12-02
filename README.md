# Local PDM/PLM Inventory Tool

A local-only CLI utility for indexing SolidWorks PDM and Aras Innovator PLM files, cross-referencing them against local drives, and running read-only PeopleSoft/Denodo queries.

## Features

- **PDM Indexing**: Scan mapped PDM drives and capture file metadata (name, path, dates)
- **PLM Scraping**: Selenium-based web scraping with configurable selectors for any PLM UI
- **Presence Checking**: Cross-reference remote PLM files against local filesystem
- **PeopleSoft Search**: Run read-only SQL queries via Denodo ODBC
- **Local Search**: Search indexed inventory by filename or path

## Prerequisites

- Python 3.10+ with `pip install -r requirements.txt`
- Chrome/Chromium browser (for PLM scraping)
- Chromedriver in `bin/` directory or on PATH
- pyodbc and configured Denodo DSN (for PeopleSoft queries)
- Windows environment (for PDM mapped drives)

## Installation

```bash
# Clone and install dependencies
pip install -r requirements.txt

# Copy and configure settings
cp config/settings.example.json config/settings.json
```

## Configuration

Edit `config/settings.json`:

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
    "start_path": "/",
    "selectors": { ... }
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

### PLM Selectors

The PLM scraper uses configurable CSS/XPath selectors. Each selector key accepts comma-separated fallbacks:

| Key | Purpose | Default |
|-----|---------|---------|
| `login_form` | Login form container | `form#login, form[name='login']` |
| `username_field` | Username input | `input[name='username'], input#username` |
| `password_field` | Password input | `input[type='password']` |
| `login_button` | Submit button | `button[type='submit'], #loginBtn` |
| `logged_in_indicator` | Element proving logged-in state | `.user-profile, .logout-btn` |
| `main_grid` | File list table/grid | `#mainGrid, .search-grid, table.aras-grid` |
| `grid_rows` | Individual file rows | `tr.grid-row, tr[data-id], tbody tr` |
| `item_name` | File name cell | `.item-name, [data-field='name']` |
| `next_page` | Pagination next button | `.next-page, a[rel='next']` |

Override any selector in your `settings.json` to match your PLM's DOM structure.

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

1. **First run**: Set `headless: false` in config to watch the browser
2. **Inspect DOM**: Use browser DevTools to find correct selectors for your PLM
3. **Update selectors**: Add site-specific selectors to `config/settings.json`
4. **Handle MFA**: On first login, complete MFA manually; use `--save-cookies` to persist session
5. **Production**: Once working, set `headless: true` for unattended runs

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
- **Selector fallbacks**: PLM selectors use comma-separated CSS selectors tried in order
- **Defensive error handling**: Stale element exceptions, timeouts, and partial failures are caught and logged

### Extending the PLM Scraper

To add support for a new PLM system:

1. Create a new selector profile in `DEFAULT_SELECTORS` or override via config
2. If the PLM uses iframes, add frame-switching logic in `_find_element()`
3. For AJAX-heavy UIs, increase `wait_timeout` and add explicit waits
4. For non-standard auth (SAML, OAuth), extend `login()` with redirect handling

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
| Support new PLM | `src/indexer/plm.py` | Update `DEFAULT_SELECTORS` |
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
