import json
from pathlib import Path
from typing import Optional, Dict, Any

class CheckpointManager:
    def __init__(self, state_file: Path):
        self.state_file = state_file
        self.state: Dict[str, Any] = {}
        self._load()

    def _load(self):
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r') as f:
                    self.state = json.load(f)
            except json.JSONDecodeError:
                self.state = {}

    def save_checkpoint(self, key: str, value: Any):
        self.state[key] = value
        with open(self.state_file, 'w') as f:
            json.dump(self.state, f)

    def get_checkpoint(self, key: str) -> Optional[Any]:
        return self.state.get(key)
    
    def clear(self):
        if self.state_file.exists():
            self.state_file.unlink()
            self.state = {}

