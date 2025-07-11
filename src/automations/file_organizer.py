"""
Example file organizer automation
"""
from pathlib import Path
from typing import Dict, List
from automations.base_automation import BaseAutomation
from utils.file_utils import ensure_directory, get_files_by_extension, move_file
from config.settings import SUPPORTED_EXTENSIONS


class FileOrganizerAutomation(BaseAutomation):
    """
    Automation to organize files by type into separate folders
    """
    
    def __init__(self, source_directory: str, target_directory: str):
        """
        Initialize file organizer
        
        Args:
            source_directory: Directory to organize files from
            target_directory: Directory to organize files into
        """
        super().__init__("file_organizer")
        self.source_dir = Path(source_directory)
        self.target_dir = Path(target_directory)
    
    def run(self) -> bool:
        """
        Organize files by type into separate folders
        
        Returns:
            True if successful, False otherwise
        """
        if not self.source_dir.exists():
            self.logger.error(f"Source directory does not exist: {self.source_dir}")
            return False
        
        # Create target directory structure
        ensure_directory(self.target_dir)
        
        organized_count = 0
        
        # Organize files by type
        for file_type, extensions in SUPPORTED_EXTENSIONS.items():
            files = get_files_by_extension(self.source_dir, extensions)
            
            if files:
                type_dir = ensure_directory(self.target_dir / file_type)
                
                for file_path in files:
                    target_path = type_dir / file_path.name
                    if move_file(file_path, target_path):
                        organized_count += 1
        
        # Handle files with unknown extensions
        remaining_files = [f for f in self.source_dir.rglob("*") if f.is_file()]
        if remaining_files:
            other_dir = ensure_directory(self.target_dir / "other")
            for file_path in remaining_files:
                target_path = other_dir / file_path.name
                if move_file(file_path, target_path):
                    organized_count += 1
        
        self.logger.info(f"Organized {organized_count} files")
        return True


# Example usage function
def organize_downloads():
    """Example function to organize downloads folder"""
    downloads_path = Path.home() / "Downloads"
    organized_path = Path.home() / "Downloads" / "Organized"
    
    organizer = FileOrganizerAutomation(str(downloads_path), str(organized_path))
    return organizer.execute()
