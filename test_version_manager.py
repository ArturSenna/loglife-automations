"""
Test script for version manager functionality
Run this to verify the update system works correctly
"""
import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from utils.version_manager import VersionManager, get_version_info

def test_version_manager():
    print("=" * 60)
    print("LogLife Operations - Version Manager Test")
    print("=" * 60)
    print()
    
    # Create version manager instance
    vm = VersionManager()
    
    # Test 1: Get version info
    print("Test 1: Get Version Info")
    print("-" * 60)
    version_info = get_version_info()
    print(f"✓ Current Version: {version_info}")
    print(f"✓ Version: {vm.current_version}")
    print(f"✓ Build Date: {vm.build_date}")
    print()
    
    # Test 2: Get local version info
    print("Test 2: Get Local Version Info")
    print("-" * 60)
    local_info = vm.get_local_version_info()
    for key, value in local_info.items():
        print(f"✓ {key}: {value}")
    print()
    
    # Test 3: Check for updates
    print("Test 3: Check for Updates")
    print("-" * 60)
    print("Checking for updates from GitHub...")
    
    has_update, remote_version, download_url, update_notes = vm.check_for_updates(show_no_update_message=False)
    
    if has_update:
        print(f"✓ Update Available!")
        print(f"  - Remote Version: {remote_version}")
        print(f"  - Download URL: {download_url}")
        print(f"  - Notes: {update_notes}")
    elif remote_version:
        print(f"✓ No Update Available")
        print(f"  - Remote Version: {remote_version}")
        print(f"  - You are up to date!")
    else:
        print("✗ Could not check for updates")
        print("  - Possible reasons:")
        print("    • No internet connection")
        print("    • GitHub API not accessible")
        print("    • version.json not found on GitHub")
    print()
    
    # Test 4: Version comparison
    print("Test 4: Version Comparison Tests")
    print("-" * 60)
    test_cases = [
        ("1.0.0", "1.0.1", True, "Patch update"),
        ("1.0.0", "1.1.0", True, "Minor update"),
        ("1.0.0", "2.0.0", True, "Major update"),
        ("1.1.0", "1.0.0", False, "Downgrade (should be False)"),
        ("1.0.0", "1.0.0", False, "Same version"),
    ]
    
    all_passed = True
    for local, remote, expected, description in test_cases:
        result = vm._compare_versions(remote, local)
        status = "✓" if result == expected else "✗"
        all_passed = all_passed and (result == expected)
        print(f"{status} {description}: {local} -> {remote} = {result}")
    
    print()
    if all_passed:
        print("✓ All version comparison tests passed!")
    else:
        print("✗ Some version comparison tests failed!")
    
    print()
    print("=" * 60)
    print("Test Complete!")
    print("=" * 60)
    
    # Interactive test
    print()
    response = input("Do you want to test the update dialog? (y/n): ")
    if response.lower() == 'y':
        print("Opening update check dialog...")
        vm.force_update_check()
    
    print()
    print("Version Manager is working correctly!")
    print("You can now build your executable with confidence.")

if __name__ == "__main__":
    try:
        test_version_manager()
    except KeyboardInterrupt:
        print("\n\nTest cancelled by user.")
    except Exception as e:
        print(f"\n✗ Error during test: {e}")
        import traceback
        traceback.print_exc()
    
    input("\nPress Enter to exit...")
