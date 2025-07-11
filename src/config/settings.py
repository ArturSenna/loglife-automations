"""
Main application configuration settings
"""
import os
from pathlib import Path

# Project root directory
PROJECT_ROOT = Path(__file__).parent.parent.parent

# Data directories
DATA_DIR = PROJECT_ROOT / "data"
LOGS_DIR = PROJECT_ROOT / "logs"

# Logging configuration
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"

# Default settings
DEFAULT_TIMEOUT = 30
MAX_RETRIES = 3

# API configurations (add your API keys as environment variables)
API_KEYS = {
    "example_api": os.getenv("EXAMPLE_API_KEY"),
    # Add more APIs as needed
}

# File extensions for different automation types
SUPPORTED_EXTENSIONS = {
    "documents": [".pdf", ".docx", ".txt", ".md"],
    "images": [".jpg", ".jpeg", ".png", ".gif", ".bmp"],
    "data": [".csv", ".json", ".xlsx", ".xml"],
}
