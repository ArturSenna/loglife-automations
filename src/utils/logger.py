"""
Logging utility for the automation system
"""
import logging
import sys
from pathlib import Path
from config.settings import LOGS_DIR, LOG_LEVEL, LOG_FORMAT


def setup_logger(name: str, log_file: str = None) -> logging.Logger:
    """
    Set up a logger with both file and console handlers
    
    Args:
        name: Logger name
        log_file: Optional log file name (will be created in logs directory)
    
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, LOG_LEVEL.upper()))
    
    # Remove existing handlers to avoid duplicates
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Create formatter
    formatter = logging.Formatter(LOG_FORMAT)
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler (if log_file is specified)
    if log_file:
        LOGS_DIR.mkdir(exist_ok=True)
        file_handler = logging.FileHandler(LOGS_DIR / log_file)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str) -> logging.Logger:
    """Get or create a logger with the given name"""
    return setup_logger(name, f"{name}.log")
