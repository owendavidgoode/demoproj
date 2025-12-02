from typing import List, Dict, Generator
from src.storage.inventory import InventoryReader
from src.utils.logging import get_logger

logger = get_logger("local_search")

class LocalSearch:
    def __init__(self, inventory: InventoryReader):
        self.inventory = inventory

    def search(self, term: str) -> Generator[Dict, None, None]:
        """
        Search the inventory for the term.
        """
        logger.info(f"Searching index for '{term}'...")
        return self.inventory.search(term)

    def report_missing(self) -> List[Dict]:
        """
        Compare PLM items to PDM items to find missing local files.
        """
        all_items = self.inventory.get_all()
        plm_items = [i for i in all_items if i.get('source') == 'plm']
        pdm_items = [i for i in all_items if i.get('source') == 'pdm']
        
        # Simple name matching for now.
        pdm_names = {i.get('name') for i in pdm_items if i.get('name')}
        
        missing = []
        for item in plm_items:
            if item.get('name') not in pdm_names:
                item_copy = item.copy()
                item_copy['status'] = 'missing_locally'
                missing.append(item_copy)
                
        return missing
