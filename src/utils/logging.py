import logging
import sys
from pathlib import Path

def setup_logging(verbose: bool = False, log_file: Path = None):
    """
    Configure logging for the application.
    
    Args:
        verbose (bool): If True, set level to DEBUG.
        log_file (Path): Optional path to write logs to.
    """
    level = logging.DEBUG if verbose else logging.INFO
    
    handlers = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file))
        
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=handlers
    )
    
    # Quiet down some 3rd party libs
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("selenium").setLevel(logging.WARNING)

def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)

