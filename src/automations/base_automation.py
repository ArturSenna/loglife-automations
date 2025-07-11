"""
Base automation class for creating consistent automation scripts
"""
from abc import ABC, abstractmethod
from typing import Any, Dict, Optional
from utils.logger import get_logger


class BaseAutomation(ABC):
    """
    Base class for all automation scripts
    """
    
    def __init__(self, name: str, config: Optional[Dict[str, Any]] = None):
        """
        Initialize base automation
        
        Args:
            name: Name of the automation
            config: Optional configuration dictionary
        """
        self.name = name
        self.config = config or {}
        self.logger = get_logger(f"automation.{name}")
        self.logger.info(f"Initialized automation: {name}")
    
    @abstractmethod
    def run(self) -> bool:
        """
        Main execution method for the automation
        
        Returns:
            True if successful, False otherwise
        """
        pass
    
    def pre_run(self) -> bool:
        """
        Pre-execution setup (override if needed)
        
        Returns:
            True if setup successful, False otherwise
        """
        self.logger.info(f"Starting pre-run setup for {self.name}")
        return True
    
    def post_run(self, success: bool) -> None:
        """
        Post-execution cleanup (override if needed)
        
        Args:
            success: Whether the main execution was successful
        """
        status = "successful" if success else "failed"
        self.logger.info(f"Completed post-run cleanup for {self.name} - {status}")
    
    def execute(self) -> bool:
        """
        Full execution workflow: pre_run -> run -> post_run
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.pre_run():
                self.logger.error(f"Pre-run setup failed for {self.name}")
                return False
            
            success = self.run()
            self.post_run(success)
            
            if success:
                self.logger.info(f"Automation {self.name} completed successfully")
            else:
                self.logger.error(f"Automation {self.name} failed")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Unexpected error in automation {self.name}: {e}")
            self.post_run(False)
            return False
