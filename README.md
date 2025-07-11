# LogLife Automations

A Python automation system for organizing and automating daily tasks.

## 📁 Project Structure

```
loglife-automations/
├── src/                          # Source code
│   ├── automations/              # Automation scripts
│   │   ├── __init__.py
│   │   ├── base_automation.py    # Base class for automations
│   │   └── file_organizer.py     # File organization automation
│   ├── config/                   # Configuration files
│   │   ├── __init__.py
│   │   └── settings.py           # Application settings
│   ├── utils/                    # Utility functions
│   │   ├── __init__.py
│   │   ├── logger.py             # Logging utilities
│   │   └── file_utils.py         # File system utilities
│   └── __init__.py
├── tests/                        # Test files
├── logs/                         # Log files (auto-generated)
├── data/                         # Data files
├── docs/                         # Documentation
├── scripts/                      # Utility scripts
├── main.py                       # Main entry point
├── requirements.txt              # Python dependencies
└── README.md                     # This file
```

## 🚀 Getting Started

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Installation

1. Clone or download this repository
2. Navigate to the project directory:
   ```bash
   cd loglife-automations
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 🐍 Setting Up Python Virtual Environment

### Windows Setup

1. **Run the setup script:**
   ```cmd
   setup_venv.bat
   ```

2. **Activate the virtual environment:**
   ```cmd
   activate_venv.bat
   ```

3. **Run automations:**
   ```cmd
   run.bat
   ```

### Linux/macOS Setup

1. **Make the script executable and run:**
   ```bash
   chmod +x setup_venv.sh
   ./setup_venv.sh
   ```

2. **Activate the virtual environment:**
   ```bash
   source venv/bin/activate
   ```

3. **Run automations:**
   ```bash
   python main.py
   ```

### Manual Setup (Alternative)

If you prefer to set up manually:

```bash
# Create virtual environment
python -m venv venv

# Activate (Windows)
venv\Scripts\activate

# Activate (Linux/macOS)
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Usage

Run the main automation script:
```bash
python main.py
```

## 🔧 Creating New Automations

1. Create a new Python file in `src/automations/`
2. Inherit from `BaseAutomation` class
3. Implement the `run()` method
4. Add your automation to `main.py`

Example:
```python
from src.automations.base_automation import BaseAutomation

class MyAutomation(BaseAutomation):
    def __init__(self):
        super().__init__("my_automation")
    
    def run(self) -> bool:
        # Your automation logic here
        self.logger.info("Running my automation")
        return True
```

## 📝 Configuration

Edit `src/config/settings.py` to modify:
- Log levels and formats
- File paths
- API keys (set as environment variables)
- Supported file extensions

## 🔍 Available Automations

### File Organizer
Organizes files by type into separate folders:
- Documents: `.pdf`, `.docx`, `.txt`, `.md`
- Images: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`
- Data: `.csv`, `.json`, `.xlsx`, `.xml`

## 📊 Logging

Logs are automatically created in the `logs/` directory. Each automation creates its own log file with timestamps and detailed information.

## 🤝 Contributing

1. Create new automations in the `src/automations/` folder
2. Add tests in the `tests/` folder
3. Update documentation as needed

## 📄 License

This project is for personal use. Feel free to modify and extend as needed.
