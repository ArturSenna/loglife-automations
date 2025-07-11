"""
Main entry point for running automations
"""
import sys
from pathlib import Path

# Add src to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from automations.file_organizer import organize_downloads
from utils.logger import get_logger

logger = get_logger("main")


def main():
    """Main function to run automations"""
    logger.info("Starting LogLife Automations")
    
    # Example: Run file organizer
    try:
        success = organize_downloads()
        if success:
            logger.info("File organization completed successfully")
        else:
            logger.error("File organization failed")
    except Exception as e:
        logger.error(f"Error running file organizer: {e}")
    
    logger.info("LogLife Automations finished")


if __name__ == "__main__":
    main()
