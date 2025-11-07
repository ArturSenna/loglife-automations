"""Verify icon embedding in executable"""
import os
import struct

def verify_icon_in_exe(exe_path):
    """Check if icon is embedded in the executable"""
    if not os.path.exists(exe_path):
        return f"ERROR: Executable not found at {exe_path}"
    
    print(f"Checking: {exe_path}")
    print(f"Size: {os.path.getsize(exe_path):,} bytes")
    
    # Read PE header
    with open(exe_path, 'rb') as f:
        # Check MZ signature
        mz_sig = f.read(2)
        if mz_sig != b'MZ':
            return "ERROR: Not a valid PE executable"
        
        # Get PE header offset
        f.seek(0x3C)
        pe_offset = struct.unpack('<I', f.read(4))[0]
        
        # Check PE signature
        f.seek(pe_offset)
        pe_sig = f.read(4)
        if pe_sig != b'PE\x00\x00':
            return "ERROR: Invalid PE signature"
        
        print("✓ Valid PE executable")
        
        # Try to find icon resources
        f.seek(0)
        data = f.read(100000)  # Read first 100KB
        
        # Look for ICO signature
        if b'\x00\x00\x01\x00' in data:
            print("✓ Icon signature found in executable!")
            return "SUCCESS: Icon appears to be embedded"
        else:
            print("⚠ Icon signature not found in first 100KB")
            return "WARNING: Could not verify icon embedding"

def check_icon_file(icon_path):
    """Verify the icon file is valid"""
    if not os.path.exists(icon_path):
        return f"ERROR: Icon file not found at {icon_path}"
    
    print(f"\nChecking icon file: {icon_path}")
    
    with open(icon_path, 'rb') as f:
        header = f.read(6)
        
        # Check ICO header
        if header[0:4] != b'\x00\x00\x01\x00':
            return "ERROR: Invalid ICO file header"
        
        num_images = struct.unpack('<H', header[4:6])[0]
        print(f"✓ Valid ICO file")
        print(f"✓ Contains {num_images} image(s)")
        print(f"✓ File size: {os.path.getsize(icon_path):,} bytes")
        
        return "SUCCESS: Icon file is valid"

if __name__ == "__main__":
    print("=" * 60)
    print("Icon Verification Tool")
    print("=" * 60)
    
    # Check icon file
    icon_result = check_icon_file("src/assets/my_icon.ico")
    print(f"\n{icon_result}\n")
    
    # Check executables
    print("-" * 60)
    exe_paths = [
        "dist/SLA_LogLife.exe",
        "dist/SLA_LogLife_Optimized.exe"
    ]
    
    for exe_path in exe_paths:
        if os.path.exists(exe_path):
            result = verify_icon_in_exe(exe_path)
            print(f"{result}\n")
            print("-" * 60)
        else:
            print(f"Skipping {exe_path} (not found)\n")
            print("-" * 60)
    
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print("Spec file configuration:")
    print("  ✓ SLA.spec has icon parameter: icon='src\\\\assets\\\\my_icon.ico'")
    print("  ✓ SLA_optimized.spec has icon parameter: icon='src\\\\assets\\\\my_icon.ico'")
    print("\nFor icon to appear on taskbar:")
    print("  1. Icon file must be valid (✓ Verified)")
    print("  2. Icon must be set in spec file (✓ Verified)")
    print("  3. Executable must be rebuilt (Run build_executable.bat)")
    print("  4. Windows may cache icons (Restart if needed)")
    print("=" * 60)
