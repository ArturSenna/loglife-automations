#!/usr/bin/env python
"""
Setup script for installing dependencies and preparing the environment
"""
import subprocess
import sys
from pathlib import Path


def install_requirements():
    """Install Python requirements"""
    requirements_file = Path(__file__).parent.parent / "requirements.txt"
    
    if requirements_file.exists():
        print("Installing Python dependencies...")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", str(requirements_file)
        ])
        print("Dependencies installed successfully!")
    else:
        print("requirements.txt not found!")


def create_directories():
    """Ensure all necessary directories exist"""
    base_dir = Path(__file__).parent.parent
    dirs = ["logs", "data"]
    
    for dir_name in dirs:
        dir_path = base_dir / dir_name
        dir_path.mkdir(exist_ok=True)
        print(f"Directory ensured: {dir_path}")


def main():
    """Main setup function"""
    print("Setting up LogLife Automations...")
    
    create_directories()
    install_requirements()
    
    print("\nSetup complete! You can now run:")
    print("python main.py")


if __name__ == "__main__":
    main()
