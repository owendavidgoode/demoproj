import json
import os
from pathlib import Path
from typing import List, Dict, Optional, Generator, Any
from src.utils.logging import get_logger

logger = get_logger("inventory")

class InventoryWriter:
    """
    Writes inventory items to a JSON file in a streaming fashion.
    Structure:
    {
        "items": [ ... ],
        "summary": { ... }
    }
    """
    def __init__(self, file_path: Path, overwrite: bool = False):
        self.file_path = file_path
        self.overwrite = overwrite
        self.file_handle = None
        self.first_item = True
        self.count = 0
        
        # Stats
        self.stats = {
            "total_pdm": 0,
            "total_plm": 0,
            "matched": 0,
            "missing_locally": 0
        }
        
        if self.file_path.exists() and not self.overwrite:
             raise FileExistsError(f"File {self.file_path} exists. Use --force to overwrite.")
        
        # Ensure dir exists
        self.file_path.parent.mkdir(parents=True, exist_ok=True)

    def __enter__(self):
        self.file_handle = open(self.file_path, 'w', encoding='utf-8')
        # Write header
        self.file_handle.write('{\n  "items": [\n')
        return self

    def add_item(self, item: Dict):
        if not self.first_item:
            self.file_handle.write(',\n')
        
        source = item.get("source")
        if source == "pdm":
            self.stats["total_pdm"] += 1
        elif source == "plm":
            self.stats["total_plm"] += 1
            if item.get("present_locally"):
                self.stats["matched"] += 1
            else:
                self.stats["missing_locally"] += 1
        
        json.dump(item, self.file_handle)
        self.first_item = False
        self.count += 1

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.file_handle:
            # Close array
            self.file_handle.write('\n  ],\n')
            
            # Write summary
            summary = {
                "total_items": self.count,
                "stats": self.stats,
                "status": "completed" if exc_type is None else "failed"
            }
            self.file_handle.write(f'  "summary": {json.dumps(summary)}\n')
            self.file_handle.write('}')
            self.file_handle.close()
            logger.info(f"Inventory saved to {self.file_path}. Total: {self.count}. Stats: {self.stats}")

class InventoryReader:
    """
    Reads inventory for search purposes.
    """
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.data: Dict[str, Any] = {}
        
    def load(self):
        if not self.file_path.exists():
            return
        
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except json.JSONDecodeError:
            logger.error(f"Failed to decode inventory at {self.file_path}")
            self.data = {"items": []}

    def search(self, term: str) -> Generator[Dict, None, None]:
        term = term.lower()
        items = self.data.get("items", [])
        for item in items:
            name = item.get('name', '').lower()
            lpath = str(item.get('local_path', '') or '').lower()
            rpath = str(item.get('remote_path', '') or '').lower()
            
            if term in name or term in lpath or term in rpath:
                yield item

    def get_all(self) -> List[Dict]:
        return self.data.get("items", [])
