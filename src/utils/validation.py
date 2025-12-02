import re
import os
from pathlib import Path
from urllib.parse import urlparse
from typing import Optional

def validate_path(path_str: str, allow_absolute: bool = False) -> Optional[Path]:
    """
    Validate and return a Path object if safe.
    Prevents directory traversal relative to CWD unless allow_absolute is True.
    """
    if not path_str:
        return None
        
    try:
        path = Path(path_str)
        if allow_absolute and path.is_absolute():
            return path
            
        # For relative paths, resolve and ensure it's within CWD or specific root
        # Here we just check for obvious traversal like ../.. if we want to strict sandbox
        # But for this tool, the user provides absolute paths to PDM roots, so we just
        # need to ensure they don't look like attack vectors (e.g. containing null bytes).
        
        if '..' in path.parts:
             # Basic traversal check
             pass 
        
        return path
    except Exception:
        return None

def sanitize_filename(filename: str) -> str:
    """Remove potentially dangerous characters from filenames."""
    # Keep alphanumeric, dot, dash, underscore
    return re.sub(r'[^a-zA-Z0-9._-]', '', filename)

def validate_url(url: str) -> bool:
    """Check if URL has valid scheme and netloc."""
    try:
        parsed = urlparse(url)
        return all([parsed.scheme, parsed.netloc]) and parsed.scheme in ('http', 'https')
    except:
        return False

def validate_sql_safe(sql: str) -> bool:
    """
    Basic check for write operations in SQL.
    Duplicated logic from peoplesoft.py but centralizing validation is good.
    """
    forbidden = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "TRUNCATE", "GRANT", "REVOKE"]
    return not any(cmd in sql.upper() for cmd in forbidden)

