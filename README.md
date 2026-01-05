# 7-Zip Windows 11 Context Menu

Add 7-Zip to Windows 11's **new context menu** (no need to click "Show more options"), organized with submenus.

## Features

When right-clicking on files or folders, a **7-Zip** menu appears with the following submenus:

### Archive Files (.7z, .zip, .rar, etc.)
- Extract Here
- Extract to Subfolder

### Regular Files/Folders
- Add to .7z
- Add to .zip

Supports automatic language switching between English and Chinese.

## System Requirements

- Windows 11 (version 21H2 or higher)
- [Visual Studio 2026 or higher ](https://visualstudio.microsoft.com/) or [Build Tools](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2026) (C++ desktop development workload required)
- [CMake](https://cmake.org/download/) 3.20+
- 7-Zip installed at the default path `C:\Program Files\7-Zip`

## Build and Installation

### 1. Enable Developer Mode

Open PowerShell as **Administrator** and run:

```powershell
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /v AllowDevelopmentWithoutDevLicense /t REG_DWORD /d 1 /f
```

Or via GUI: Settings > Privacy & Security > Developer Options > Enable "Developer Mode"

### 2. Build

Open **"Developer Command Prompt for VS 2026"** or **PowerShell**:

```powershell
# Configure
cmake -B build -G "Visual Studio 18 2026" -A x64

# Build
cmake --build build --config Release

# Install to install directory
cmake --install build --config Release
```

### 3. Register

Open PowerShell as **Administrator**:

```powershell
cd build/install
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

### 4. Restart Explorer

```powershell
Stop-Process -Name explorer -Force; Start-Process explorer
```

Or sign out and sign back in.

## Uninstallation

Open PowerShell as **Administrator**:

```powershell
cd build/install
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## Project Structure

```
7z/
├── CMakeLists.txt          # CMake configuration
├── README.md
├── src/
│   ├── ContextMenu.cpp     # IExplorerCommand COM implementation
│   ├── ContextMenu.h       # Header file definitions
│   └── ContextMenu.def     # DLL exports
├── scripts/
│   ├── install.ps1         # Installation script
│   └── uninstall.ps1       # Uninstallation script
├── Package/                # Source resource files
│   ├── AppxManifest.xml    # MSIX Sparse Package manifest
│   └── Assets/             # Icons directory
└── build/                  # Build directory
    └── install/            # Installation output
        ├── 7ZipContextMenu.dll
        ├── 7ZipContextMenu.pdb
        ├── AppxManifest.xml
        ├── Assets/
        ├── passwords.txt   # Password storage file
        ├── install.ps1
        └── uninstall.ps1
```

## Custom 7-Zip Path

If 7-Zip is installed in a different location, modify the paths in `src/ContextMenu.cpp`:

```cpp
static const wchar_t* g_7zPath = L"C:\\Program Files\\7-Zip";
static const wchar_t* g_7zDll = L"C:\\Program Files\\7-Zip\\7z.dll";
static const wchar_t* g_7zFM = L"C:\\Program Files\\7-Zip\\7zFM.exe";
```

Then rebuild and reinstall.

## Troubleshooting

### Package Registration Failed (0x80073CFF)
- Ensure Developer Mode is enabled (see step 1)
- Ensure install.ps1 is run as Administrator

### Menu Not Appearing
- Restart Windows Explorer: `Stop-Process -Name explorer -Force; Start-Process explorer`
- Or sign out and sign back in

### Build Errors
- Ensure Visual Studio 2026's C++ desktop development workload is installed
- Ensure CMake is added to PATH

### PowerShell Script Cannot Run
Use the `-ExecutionPolicy Bypass` parameter:
```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

## Technical Details

This project uses:
- **COM Shell Extension**: Implements `IExplorerCommand` interface with submenu support
- **MSIX Sparse Package**: Allows non-store apps to register in Windows 11's new context menu
- **7-Zip SDK**: Directly calls 7z.dll for compression/decompression
- **Self-signed Certificate**: Automatically created and installed to trusted root certificates

## License

MIT License
