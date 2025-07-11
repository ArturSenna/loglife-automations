# Example test for file organizer automation
import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path
import sys

# Add src to Python path for testing
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from src.automations.file_organizer import FileOrganizerAutomation


class TestFileOrganizerAutomation(unittest.TestCase):
    """Test cases for FileOrganizerAutomation"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.source_dir = "/test/source"
        self.target_dir = "/test/target"
        self.organizer = FileOrganizerAutomation(self.source_dir, self.target_dir)
    
    def test_initialization(self):
        """Test proper initialization"""
        self.assertEqual(self.organizer.name, "file_organizer")
        self.assertEqual(str(self.organizer.source_dir), self.source_dir)
        self.assertEqual(str(self.organizer.target_dir), self.target_dir)
    
    @patch('src.automations.file_organizer.Path.exists')
    def test_run_source_not_exists(self, mock_exists):
        """Test run method when source directory doesn't exist"""
        mock_exists.return_value = False
        result = self.organizer.run()
        self.assertFalse(result)
    
    # Add more test cases as needed


if __name__ == '__main__':
    unittest.main()
