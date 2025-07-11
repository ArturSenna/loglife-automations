"""
File system utilities for automation tasks
"""
import os
import shutil
from pathlib import Path
from typing import List, Optional
from utils.logger import get_logger

logger = get_logger(__name__)


def ensure_directory(path: Path) -> Path:
    """
    Ensure a directory exists, create if it doesn't
    
    Args:
        path: Directory path
    
    Returns:
        Path object of the directory
    """
    path.mkdir(parents=True, exist_ok=True)
    logger.info(f"Directory ensured: {path}")
    return path


def get_files_by_extension(directory: Path, extensions: List[str]) -> List[Path]:
    """
    Get all files in directory with specified extensions
    
    Args:
        directory: Directory to search
        extensions: List of file extensions (e.g., ['.txt', '.pdf'])
    
    Returns:
        List of file paths
    """
    files = []
    for ext in extensions:
        files.extend(directory.glob(f"**/*{ext}"))
    
    logger.info(f"Found {len(files)} files with extensions {extensions} in {directory}")
    return files


def move_file(source: Path, destination: Path) -> bool:
    """
    Move file from source to destination
    
    Args:
        source: Source file path
        destination: Destination file path
    
    Returns:
        True if successful, False otherwise
    """
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(source), str(destination))
        logger.info(f"Moved file: {source} -> {destination}")
        return True
    except Exception as e:
        logger.error(f"Failed to move file {source} to {destination}: {e}")
        return False


def copy_file(source: Path, destination: Path) -> bool:
    """
    Copy file from source to destination
    
    Args:
        source: Source file path
        destination: Destination file path
    
    Returns:
        True if successful, False otherwise
    """
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(source), str(destination))
        logger.info(f"Copied file: {source} -> {destination}")
        return True
    except Exception as e:
        logger.error(f"Failed to copy file {source} to {destination}: {e}")
        return False


def safe_filename(filename: str) -> str:
    """
    Convert a string to a safe filename by removing/replacing invalid characters
    
    Args:
        filename: Original filename
    
    Returns:
        Safe filename string
    """
    invalid_chars = '<>:"/\\|?*'
    safe_name = filename
    for char in invalid_chars:
        safe_name = safe_name.replace(char, '_')
    
    return safe_name.strip()
