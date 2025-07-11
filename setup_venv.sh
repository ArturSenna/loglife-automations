#!/bin/bash
# Setup script for creating Python virtual environment on Unix/Linux/macOS

echo "Setting up Python virtual environment for LogLife Automations..."

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed. Please install Python 3 and try again."
    exit 1
fi

# Display Python version
echo "Current Python version:"
python3 --version

# Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv

# Activate virtual environment and install dependencies
echo "Activating virtual environment..."
source venv/bin/activate

echo "Upgrading pip..."
python -m pip install --upgrade pip

echo "Installing project dependencies..."
pip install -r requirements.txt

echo ""
echo "==================================="
echo "Virtual environment setup complete!"
echo "==================================="
echo ""
echo "To activate the virtual environment, run:"
echo "  source venv/bin/activate"
echo ""
echo "To deactivate, run:"
echo "  deactivate"
echo ""
echo "To run the automation system:"
echo "  python main.py"
echo ""
