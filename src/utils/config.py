import json
import os
from pathlib import Path
from typing import Dict, Any, Optional

class ConfigError(Exception):
    pass

class Config:
    def __init__(self, config_path: Optional[Path] = None):
        self._config: Dict[str, Any] = {}
        self.load(config_path)

    def load(self, config_path: Optional[Path] = None):
        """Load configuration from a JSON file."""
        if config_path is None:
            # Default to config/settings.json in the project root if not provided
            # Assuming run from root or src handling
            base_dir = Path(os.getcwd())
            config_path = base_dir / "config" / "settings.json"

        if not config_path.exists():
            # If explicit path provided and missing, error. If default, warn or empty?
            # For now, we'll initialize empty if default is missing, but log warning in real usage
            # Or simpler: just let it be empty dict or raise if essential
            print(f"Warning: Config file not found at {config_path}")
            return

        try:
            with open(config_path, 'r') as f:
                self._config = json.load(f)
        except json.JSONDecodeError as e:
            raise ConfigError(f"Failed to parse config file: {e}")

    def get(self, key: str, default: Any = None) -> Any:
        """Retrieve a value by dot-notation key (e.g., 'pdm.roots')."""
        keys = key.split('.')
        value = self._config
        for k in keys:
            if isinstance(value, dict):
                value = value.get(k)
            else:
                return default
            if value is None:
                return default
        return value

# Singleton instance placeholder if needed, or instantiate in main

