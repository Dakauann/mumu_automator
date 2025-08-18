#!/usr/bin/env python3
"""
Test script to verify the new MuMu path detection functionality.
"""

import os
import sys

# Add the current directory to the path so we can import from main.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import find_mumu_manager_exe, get_mumu_base_path_from_manager

def test_path_detection():
    """Test the path detection with various scenarios."""
    
    print("🔍 Testing MuMu Player path detection...")
    print("-" * 50)
    
    # Test cases for different folder structures
    test_paths = [
        # New structure
        "D:\\Program Files\\Netease\\MuMuPlayerGlobal-12.0",
        "C:\\Program Files\\Netease\\MuMuPlayerGlobal-12.0",
        
        # Old structure (hypothetical)
        "D:\\Program Files\\Netease\\MuMuPlayer-12.0",
        "C:\\Program Files\\Netease\\MuMuPlayer-12.0",
        
        # Invalid paths
        "D:\\NonExistent\\Path",
        "C:\\Program Files\\SomeOtherApp",
    ]
    
    for test_path in test_paths:
        print(f"\n📂 Testing path: {test_path}")
        
        if os.path.exists(test_path):
            manager_exe = find_mumu_manager_exe(test_path)
            if manager_exe:
                print(f"   ✅ Found: {manager_exe}")
                
                # Test the reverse function
                base_path = get_mumu_base_path_from_manager(manager_exe)
                print(f"   🔄 Base path: {base_path}")
                
                if base_path == test_path:
                    print(f"   ✅ Path round-trip successful!")
                else:
                    print(f"   ❌ Path round-trip failed! Expected: {test_path}, Got: {base_path}")
            else:
                print(f"   ❌ MuMuManager.exe not found")
        else:
            print(f"   ⚠️  Path does not exist (skipping)")
    
    print("\n" + "=" * 50)
    print("🏁 Path detection test completed!")

if __name__ == "__main__":
    test_path_detection()
