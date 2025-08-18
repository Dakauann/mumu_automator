# MuMu Player Global 12.0 Folder Architecture Update

## Summary

This update modifies the MuMu Automator to work with the new folder architecture introduced in MuMu Player Global 12.0, while maintaining backwards compatibility with older versions.

## Key Changes Made

### 1. New Helper Functions Added

#### `find_mumu_manager_exe(mumu_base_path)`

- **Purpose**: Automatically detects the correct location of `MuMuManager.exe`
- **Supports**: Both new (`nx_main/`) and old (`shell/`) folder structures
- **Behavior**:
  - First tries the new structure: `nx_main/MuMuManager.exe`
  - Falls back to old structure: `shell/MuMuManager.exe`
  - Returns `None` if neither is found

#### `get_mumu_base_path_from_manager(mumu_manager_path)`

- **Purpose**: Extracts the base MuMu installation path from the manager executable path
- **Supports**: Both folder structures
- **Used**: For reverse path calculation in various functions

### 2. Updated Functions

All functions that previously hardcoded the `shell/` path have been updated:

- `discover_vm_range()` - Now uses dynamic path detection
- `get_vm_info()` - Uses `find_mumu_manager_exe()` instead of hardcoded path
- `padronize_vm_names()` - Updated for new architecture
- `create_path_config_ui()` - Enhanced validation and messaging
- `run_management()` - All VM info calls updated to use correct base path
- `main()` - Uses new path detection throughout

### 3. Folder Structure Support

#### New Structure (MuMu Player Global 12.0+)

```
D:/Program Files/Netease/MuMuPlayerGlobal-12.0/
├── nx_main/
│   ├── MuMuManager.exe          ← New location
│   └── ... (other executables)
├── nx_device/
│   └── 12.0/
│       ├── shell/
│       │   ├── MuMuNxDevice.exe
│       │   └── ... (device executables)
│       └── ... (other folders)
└── vms/
    ├── MuMuPlayerGlobal-12.0-0/
    ├── MuMuPlayerGlobal-12.0-1/
    └── ... (VM instances)
```

#### Old Structure (Previous versions)

```
D:/Program Files/Netease/MuMuPlayer-12.0/
├── shell/
│   ├── MuMuManager.exe          ← Old location
│   └── ... (other executables)
└── vms/
    ├── MuMuPlayer-12.0-0/
    └── ... (VM instances)
```

### 4. Enhanced Error Handling

- Better error messages when `MuMuManager.exe` is not found
- Clearer feedback about which folder structure is being used
- Improved path validation in the UI

### 5. UI Improvements

- Updated example paths in the configuration dialog
- Added informational text about automatic structure detection
- Better visual feedback during path validation

## Testing Results

✅ **Successfully tested with MuMu Player Global 12.0**

- Correctly detected `nx_main/MuMuManager.exe`
- Successfully discovered 14 VMs
- VM control operations working properly
- Backwards compatibility maintained

## Usage

The application now works seamlessly with both folder structures:

1. **For new installations**: Point to `D:\Program Files\Netease\MuMuPlayerGlobal-12.0`
2. **For old installations**: Point to `D:\Program Files\Netease\MuMuPlayer-12.0`

The application will automatically detect which structure is being used and adapt accordingly.

## Migration Notes

- **No user action required**: The application automatically detects the correct structure
- **Saved paths**: Previously saved paths will be re-validated automatically
- **Multiple installations**: Can switch between old and new installations seamlessly

## Developer Notes

- All hardcoded `shell/` paths have been replaced with dynamic detection
- The `find_mumu_manager_exe()` function is the central point for path resolution
- Error handling includes specific messages for missing executables
- Logging includes indicators for which structure is detected (✓ symbols)

---

_Updated for MuMu Player Global 12.0 folder architecture compatibility_
