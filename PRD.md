# PRD: Local PDM/PLM File Inventory Tool

## Overview
Local-only utility with three core functions:
- Index SolidWorks PDM (desktop client mapped as a drive) and Aras Innovator (Anovia PLM in browser)
  by signing in with your normal credentials, listing accessible files, and producing a JSON
  inventory (name, remote path, metadata), then cross-referencing designated local roots to flag
  which files are present or missing.
- PeopleSoft search function that runs user-provided SQL queries against PeopleSoft tables to return
  matching rows.
- Local/Windows file search (post-indexing) to find files by name/path across specified drives.

## Goals
- Reuse interactive login/session (no API tokens) to browse accessible files and folders.
- Emit a JSON report with remote file details and local presence status.
- Keep scope tight for single-user, read-only usage; minimal hardening.
- Provide quick local search across indexed data and Windows/local drives.

## Non-goals
- No file uploads/edits/deletes.
- No multi-user support or server deployment.
- No long-term secret storage.
- No writes to PeopleSoft or PDM/PLM.

## Users
- Single internal user (you) running locally via CLI on workstation.

## Assumptions
- PDM is accessed via desktop client mapped to a drive; files can be enumerated directly from that
  mapped drive for indexing.
- PLM is accessed in-browser with form-based login; MFA on first login, then persistent session
  allowed. Must scrape the web UI (no APIs). Playwright is unavailable; Selenium is assumed allowed
  if prepackaged.
- Windows environment (local/Windows searches run natively); Python preferred (existing
  `That_Search_Tool.py`), with Node/Go only if they fit SpiderIDE packaging.
- Remote PLM paths map to case-sensitive archive drive paths; mapping rules may need configuration.
- Corporate policy blocks installs; solution must be fully prepackaged/portable (no new system-wide
  drivers or package installs). Risk remains if headless browser binaries cannot be bundled.

## Tech Constraints & Options
- No Playwright. Automation must be SpiderIDE-friendly and prepackaged (no interactive installs).
- Selenium is acceptable if the driver/browser can be bundled portably.
- Candidate stacks (ranked):
  - Python + Selenium (or undetected-chromedriver) + BeautifulSoup for scraping, reusing existing
    Python/PyQt code where helpful.
  - Python + raw requests + HTML parsing if authentication can be captured via session cookies.
  - Go + chromedp/rod as a fallback if SpiderIDE supports bundling a headless browser driver.
- Install restrictions: if no browser driver can be bundled/approved, scraping may be blocked; may
  need to rely on existing local browser session cookies or an embedded HTML session replay approach.

## Functional Requirements
- Core functions
- PDM/PLM indexing: authenticate (PLM), enumerate files/folders (PDM via mapped drive; PLM via
    scrape), capture metadata, produce JSON inventory and presence flags.
- PeopleSoft search: prompt for SQL or accept a query file, execute read-only queries against
    PeopleSoft tables (via Denodo in current environment), return results locally (no writes).
  - Local/Windows file search: search by filename/path across configured drives or folders; leverage
    prior index where possible.
- Auth/session
  - PLM: establish session using username/password with MFA on first run; cache/reuse session cookies
    for subsequent runs when available.
  - PDM: relies on desktop-mapped drive access; no web login required for file enumeration.
  - Do not persist credentials to disk; keep in memory only.
  - Respect vault/repository selection if applicable.
- Remote listing
  - PDM: enumerate files via mapped drive; capture metadata from filesystem (name, path, created_at,
    modified_at).
  - PLM: enumerate accessible folders/files via web UI scraping/headless browser; handle pagination
    or lazy-loaded views. Use human-like request pacing (randomized delays 1-5s between pages,
    longer pauses every N requests) to avoid triggering rate limits or security alerts.
  - Capture metadata: name, remote path, created_at, modified_at, remote id if visible.
  - Apply client-side filters in-tool (folder prefix, extension, date range); default to include
    everything.
- Local scan
  - Recursively scan one or more user-specified local roots (including the PDM mapped drive and
    other Windows archive drive paths).
  - Normalize paths for comparison: case-insensitive matching for Windows paths, case-sensitive for
    Unix/archive paths. Configurable per-root if mixed environments exist.
- Output/report
  - Produce JSON array: `{ name, remote_path, remote_id?, created_at?, modified_at?,
    present_locally: boolean, local_path?: string }`.
  - Include summary stats (remote total, matches, missing).
  - Prompt before overwriting existing output files; use `--force` flag to skip confirmation.
- PeopleSoft search
  - Accept connection details/credentials at runtime (not persisted).
  - Run user-provided read-only SQL (Denodo gateway for PeopleSoft); return tabular results
    (JSON/CSV) locally.
  - Basic safety checks to prevent accidental writes (reject INSERT/UPDATE/DELETE).
- Local/Windows file search
  - Search across specified drives/folders for filenames/paths; optionally use indexed metadata for
    faster results.
  - Support simple filters (extension, path prefix).
- CLI UX
  - Flags/prompts for: server URL, vault/repo, username (password hidden), local roots, filters,
    output path, dry-run, PeopleSoft connection/query inputs, search targets.
  - Validate all inputs: sanitize paths to prevent traversal, validate URLs, reject malformed queries.
  - Client-side filtering in-app; progress indicators for remote listing and local scan; friendly
    errors.
- Logging
  - Verbose/debug flag to write request metadata (no passwords) to local log file if enabled.

## Data Handling
- Credentials never written to disk.
- Optional reuse of existing authenticated session cookies if accessible.
- JSON report stored locally; no external network beyond PDM/PLM endpoints.

## Performance
- Vault size ~20TB+; runs may be long. Stream/paginate remote listing; avoid loading full sets into
  memory. Local scans should be incremental or chunked to manage memory.
- Checkpoint/resume support: periodically write progress state (last processed path, cursor position)
  to a state file; on restart, detect incomplete runs and offer to resume from last checkpoint.
  Stream results to output file incrementally rather than buffering in memory.

## Error Handling
- Clear messages for auth failure, network timeout, permission issues, unexpected HTML or endpoint
  changes.
- Allow partial results with warnings when remote listing fails mid-run; clearly mark incomplete
    sections in output (e.g., `"status": "partial"`, `"failed_paths": [...]`).

## Extensibility (Future)
- Export CSV; incremental diff between runs; optional download of missing files.
- Broader PeopleSoft automation beyond read-only queries.

## Open Questions
1) PDM/PLM login details: target URLs for login and vault selection; where session cookies are stored
   for reuse; MFA method (TOTP/SMS) so we can automate first-login capture.
2) PDM/PLM UI structure: any known pages that list files in bulk (e.g., folder tree views) we should
   target first?
3) Path mapping: what is the root local archive path to map remote paths onto?
4) PeopleSoft/Denodo access: DB type/driver (Denodo JDBC/ODBC/REST), endpoint, and how credentials
   will be provided (prompt vs env)? Any query timeout/throttling?
5) Windows search: which drives/paths are in scope for native Windows execution (network shares,
   mapped drives, local disks)?
6) SpiderIDE packaging: which browser/driver combo is permitted (e.g., Selenium with packaged
   Chrome/Edge/IE-compat)? Any restrictions on bundling external binaries like chromedriver? If no
   installs are allowed, can we ship a portable driver/browser inside the app bundle?
