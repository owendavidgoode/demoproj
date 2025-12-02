import getpass
import argparse
import sys
import logging
import json
from pathlib import Path
from datetime import datetime

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, **kwargs): return iterable

from src.utils.logging import setup_logging, get_logger
from src.utils.config import Config
from src.indexer.pdm import PDMIndexer
from src.indexer.plm import PLMIndexer
from src.storage.inventory import InventoryReader, InventoryWriter
from src.storage.checkpoint import CheckpointManager
from src.search.peoplesoft import PeopleSoftSearch
from src.search.local import LocalSearch

logger = get_logger("cli")

def get_credentials(config, service_name):
    """
    Get credentials for a service.
    First check config (non-production/dev only), then prompt.
    """
    conf = config.get(service_name, {})
    username = conf.get("username")
    password = conf.get("password")
    
    if not username:
        username = input(f"Enter {service_name.upper()} Username: ")
    
    if not password:
        password = getpass.getpass(f"Enter {service_name.upper()} Password: ")
        
    return username, password

def cmd_index(args, config):
    if args.dry_run:
        logger.info("[DRY RUN] Starting indexing process (no changes will be saved)...")
    else:
        logger.info("Starting indexing process...")
    
    inventory_path = Path(config.get("output.path", "inventory.json"))
    checkpoint_path = Path("config/checkpoint.json")
    checkpoint_mgr = CheckpointManager(checkpoint_path)
    
    # Parse date filters
    date_from = datetime.fromisoformat(args.date_from) if args.date_from else None
    date_to = datetime.fromisoformat(args.date_to) if args.date_to else None
    
    # Filtering logic closure
    def apply_filters(item):
        if args.ext and not item.get('name', '').endswith(args.ext):
            return False
            
        if args.path_prefix:
            # Check both local and remote paths
            lpath = item.get('local_path', '') or ''
            rpath = item.get('remote_path', '') or ''
            if not (str(lpath).startswith(args.path_prefix) or str(rpath).startswith(args.path_prefix)):
                return False

        # Date filtering (modified_at preferred, created_at fallback)
        item_date_str = item.get('modified_at') or item.get('created_at')
        if item_date_str:
            try:
                item_date = datetime.fromisoformat(item_date_str)
                if date_from and item_date < date_from:
                    return False
                if date_to and item_date > date_to:
                    return False
            except ValueError:
                pass # Ignore bad dates
        return True

    # Presence map (normalized relative path -> boolean)
    # We map normalized local names/paths to check against remote
    local_presence_map = {} 
    
    # Per-root case sensitivity config
    pdm_roots = config.get("pdm.roots", [])
    # Default is case-insensitive (False) for Windows
    # In a real scenario, this map would be built from config settings
    case_sensitive_roots = {} 
    
    # Helper for case handling
    def normalize_key(path_str, root_path=None):
        if not path_str: return ""
        # If we knew the root of this path, we could lookup case sensitivity
        # For now, default to insensitive (lower)
        return path_str.lower()
    
    try:
        # Check overwrite
        if inventory_path.exists() and not args.force and not args.resume and not args.dry_run:
             logger.error(f"Output file {inventory_path} exists. Use --force to overwrite or --resume to continue.")
             sys.exit(1)
        
        # Dry Run Mock Writer
        class DryRunWriter:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def add_item(self, item): pass

        WriterClass = DryRunWriter if args.dry_run else InventoryWriter
        
        with WriterClass(inventory_path, overwrite=args.force or args.resume) as writer:
            
            # PDM Indexing
            if not args.plm_only:
                if pdm_roots:
                    logger.info("Scanning PDM...")
                    pdm_indexer = PDMIndexer(pdm_roots)
                    pbar = tqdm(desc="PDM Files", unit="file")
                    
                    for item in pdm_indexer.scan():
                        if not apply_filters(item):
                            continue
                            
                        item['source'] = 'pdm'
                        writer.add_item(item)
                        
                        # Store for presence check
                        # Use relative_path if available (preferred)
                        rel_path = item.get('relative_path')
                        if rel_path:
                             local_presence_map[rel_path.lower()] = True
                        else:
                             # Fallback to name
                             name = item.get('name') 
                             if name:
                                local_presence_map[name.lower()] = True

                        pbar.update(1)
                    pbar.close()
                    if not args.dry_run:
                        checkpoint_mgr.save_checkpoint("pdm_done", True)
                else:
                    logger.warning("No PDM roots configured.")

            # PLM Indexing
            if not args.pdm_only:
                plm_config = config.get("plm", {})
                if plm_config.get("url"):
                    # Credentials prompt
                    user, pwd = get_credentials(config, "plm")
                    # Update config in memory only
                    plm_config["username"] = user
                    plm_config["password"] = pwd
                    plm_config["save_cookies"] = args.save_cookies
                    
                    plm_indexer = None
                    try:
                        plm_indexer = PLMIndexer(plm_config)
                        
                        logger.info("Scanning PLM...")
                        pbar = tqdm(desc="PLM Items", unit="item")
                        
                        last_plm_id = checkpoint_mgr.get_checkpoint("last_plm_id")
                        
                        count = 0
                        for item in plm_indexer.scan():
                            # Resume logic could go here if item has monotonic ID
                            
                            if not apply_filters(item):
                                continue
                            
                            item['source'] = 'plm'
                            
                            # Presence check
                            # Logic: Try to match remote_path suffix with local relative path
                            # Or just match names if paths don't align
                            # For now, strict name/relative path matching
                            
                            # Mock PLM path: /Vault/Projects/A
                            # PDM relative path: Projects/A
                            # We might need to try different suffix matches or normalize
                            
                            present = False
                            
                            # Try 1: Exact relative path match (normalized)
                            # Assuming item['remote_path'] is full path
                            rpath = item.get('remote_path', '')
                            # Strip leading slash for comparison
                            if rpath.startswith('/'): rpath = rpath[1:]
                            if rpath.lower() in local_presence_map:
                                present = True
                                
                            # Try 2: Name match fallback
                            if not present:
                                name = item.get('name', '').lower()
                                if name in local_presence_map:
                                    present = True

                            item['present_locally'] = present
                            writer.add_item(item)
                            pbar.update(1)
                            
                            count += 1
                            if not args.dry_run and count % 50 == 0:
                                checkpoint_mgr.save_checkpoint("last_plm_id", item.get('remote_id'))
                                
                        pbar.close()
                    except Exception as e:
                        logger.error(f"PLM Indexing failed: {e}")
                    finally:
                        if plm_indexer:
                            plm_indexer.close()
                else:
                    logger.warning("No PLM URL configured.")
                    
    except Exception as e:
        logger.error(f"Indexing process failed: {e}")
        sys.exit(1)

    logger.info("Indexing complete.")
    if not args.dry_run:
        checkpoint_mgr.clear()

def cmd_search_ps(args, config):
    logger.info("Starting PeopleSoft search...")
    
    ps_config = config.get("peoplesoft", {})
    # Secure credential prompt if not in connection string
    # Assuming connection_string might lack pwd, but pyodbc usually needs it inline or DSN
    # If using DSN, creds might be in ODBC manager. If not, we might need to construct the string.
    # For this MVP, we rely on the connection string in config but warn if insecure?
    # Or we allow providing user/pass args to inject into DSN.
    
    # Simple approach: If config has placeholders or user wants prompt
    if args.prompt_creds:
        user, pwd = get_credentials(config, "peoplesoft")
        # Inject into connection string (naive replace or append)
        conn_str = ps_config.get("connection_string", "")
        if "PWD=" not in conn_str.upper():
             conn_str += f";UID={user};PWD={pwd}"
    else:
        conn_str = ps_config.get("connection_string")

    timeout = ps_config.get("query_timeout", 30)
    
    if not conn_str:
        logger.error("No PeopleSoft connection string configured.")
        return

    searcher = PeopleSoftSearch(conn_str, timeout)
    try:
        results = searcher.execute_query_from_file(args.query_file)
        print(json.dumps(results, indent=2, default=str))
    except Exception as e:
        logger.error(f"Search failed: {e}")

def cmd_search_local(args, config):
    logger.info(f"Starting local search for '{args.term}'...")
    
    inventory_path = Path(config.get("output.path", "inventory.json"))
    inventory = InventoryReader(inventory_path)
    inventory.load() 
    
    searcher = LocalSearch(inventory)
    
    results = list(searcher.search(args.term))
    
    if results:
        print(f"Found {len(results)} matches:")
        for item in results:
            print(f" - {item.get('name')} ({item.get('source')}): {item.get('local_path') or item.get('remote_path')}")
            if item.get('present_locally') is not None:
                 print(f"   [Present Locally: {item.get('present_locally')}]")
    else:
        print("No matches found.")

def main():
    parser = argparse.ArgumentParser(description="Local PDM/PLM Inventory Tool")
    parser.add_argument("--config", type=Path, help="Path to settings.json")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging")
    
    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    
    # Index Command
    index_parser = subparsers.add_parser("index", help="Index PDM/PLM files")
    index_parser.add_argument("--pdm-only", action="store_true", help="Index only PDM/Filesystem")
    index_parser.add_argument("--plm-only", action="store_true", help="Index only PLM/Web")
    index_parser.add_argument("--force", "-f", action="store_true", help="Overwrite existing inventory file")
    index_parser.add_argument("--resume", "-r", action="store_true", help="Resume from checkpoint (overwrite output)")
    index_parser.add_argument("--dry-run", action="store_true", help="Simulate run without writing to disk")
    index_parser.add_argument("--ext", type=str, help="Filter by file extension (e.g., .sldprt)")
    index_parser.add_argument("--path-prefix", type=str, help="Filter by path prefix")
    index_parser.add_argument("--date-from", type=str, help="Filter from date (ISO format)")
    index_parser.add_argument("--date-to", type=str, help="Filter to date (ISO format)")
    index_parser.add_argument("--save-cookies", action="store_true", help="Save PLM session cookies to disk for reuse")
    
    # Search PeopleSoft Command
    ps_parser = subparsers.add_parser("search-ps", help="Search PeopleSoft")
    ps_parser.add_argument("query_file", type=Path, help="Path to SQL query file")
    ps_parser.add_argument("--prompt-creds", action="store_true", help="Prompt for DB credentials")
    
    # Search Local Command
    local_parser = subparsers.add_parser("search-local", help="Search local index/files")
    local_parser.add_argument("term", type=str, help="Search term (filename or path)")
    
    args = parser.parse_args()
    
    setup_logging(verbose=args.verbose)
    
    try:
        config = Config(args.config)
    except Exception as e:
        logger.error(f"Configuration error: {e}")
        sys.exit(1)

    if args.command == "index":
        cmd_index(args, config)
    elif args.command == "search-ps":
        cmd_search_ps(args, config)
    elif args.command == "search-local":
        cmd_search_local(args, config)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
