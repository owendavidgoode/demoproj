import os
import datetime
from pathlib import Path
from typing import List, Dict, Generator
from src.utils.logging import get_logger

logger = get_logger("pdm_indexer")

class PDMIndexer:
    def __init__(self, roots: List[str]):
        """
        Initialize PDM Indexer.
        
        Args:
            roots: List of directory paths to scan.
        """
        self.roots = [Path(r) for r in roots]

    def scan(self) -> Generator[Dict, None, None]:
        """
        Generator that yields file metadata for every file found in roots.
        """
        for root in self.roots:
            if not root.exists():
                logger.warning(f"Root path does not exist: {root}")
                continue
            
            logger.info(f"Scanning root: {root}")
            for dirpath, _, filenames in os.walk(root):
                try:
                    # Calculate relative path from the scanned root.
                    # e.g. if root is Z:\Vault, and file is Z:\Vault\Project\Part.prt
                    # relative is Project\Part.prt
                    # This allows us to map to PLM structure if we assume PLM structure 
                    # mirrors the folder structure under the root.
                    rel_dir = Path(dirpath).relative_to(root)
                except ValueError:
                    rel_dir = Path(".")

                for f in filenames:
                    full_path = Path(dirpath) / f
                    try:
                        stat = full_path.stat()
                        
                        # Store normalized relative path (forward slashes) for comparison
                        # "Project/Part.prt"
                        rel_path_str = (rel_dir / f).as_posix()
                        
                        yield {
                            "name": f,
                            "local_path": str(full_path),
                            "relative_path": rel_path_str, # Key for presence check
                            "remote_path": None, 
                            "size": stat.st_size,
                            "modified_at": datetime.datetime.fromtimestamp(stat.st_mtime).isoformat(),
                            "created_at": datetime.datetime.fromtimestamp(stat.st_ctime).isoformat(),
                            "source": "pdm",
                            "present_locally": True,
                            "root_path": str(root)
                        }
                    except OSError as e:
                        logger.error(f"Error accessing {full_path}: {e}")
                        continue
