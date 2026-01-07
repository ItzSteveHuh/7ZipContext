# 7-Zip Windows 11 Context Menu

Add 7-Zip to Windows 11's **new context menu**, organized with submenus — no need to click "Show more options."

## Features

When right-clicking on files or folders, a **7-Zip** menu appears with the following options:

### For Archive Files (.7z, .zip, .rar, .tar, .gz, .bz2, .xz, .iso, .cab, .arj, .lzh, .tgz, .tbz2, .wim)
- **Open Archive** - Open in 7-Zip File Manager
- **Extract Here** - Extract to current directory
- **Extract to Subfolder** - Extract to a subfolder with archive name
- **Extract to Custom Folder** - Choose custom extraction location
- **Test Archive** - Verify archive integrity
- **Add to Archive** - Launch the 7-Zip add to archive GUI with custom compression options

### For Regular Files/Folders
- **Add to .7z** - Create a .7z archive
- **Add to .zip** - Create a .zip archive
- **Compress and email** - Create a .zip in the temp folder for emailing

### Additional Features
- **Password support** - Automatically handles encrypted archives with password prompts
- **Password history** - Remembers frequently used passwords for quicker access
- **Multilingual** - English and Chinese (auto-detected based on system locale)
- **Automatic detection** - Detects both 7-Zip and 7-Zip Zstandard installations

## System Requirements

- Windows 11 (version 21H2 or higher)
- [Visual Studio 2026 or higher](https://visualstudio.microsoft.com/) or [Build Tools](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2026) (C++ desktop development workload required)
- [CMake](https://cmake.org/download/) 3.20+
- 7-Zip or 7-Zip Zstandard installed (auto-detects at registry paths or common installation locations)

### Duplicate Classic Context Menu Entries
- If you see duplicate 7-Zip entries in the classic “Show more options” menu, disable 7-Zip’s own shell integration to avoid overlap with this extension.
- Recommended: open 7-Zip (or 7-Zip Zstandard) as Administrator → Tools → Options → select the 7-Zip (ZS) tab, then uncheck:
    - “Integrate 7-Zip(ZS) to shell context menu”
    - “Integrate 7-Zip(ZS) to shell context menu (32-bit)”
    - (Optional) “Cascaded context menu”
- This keeps the Windows 11 new context menu clean while preventing duplicate classic menu entries.

## Build and Installation

### 1. Enable Developer Mode

Open PowerShell as **Administrator** and run:

```powershell
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /v AllowDevelopmentWithoutDevLicense /t REG_DWORD /d 1 /f
```

Or via GUI: Settings > Privacy & security > For developers > Enable "Developer Mode"

### 2. Build

Open **"Developer Command Prompt for VS 2026"** or **PowerShell**:

```powershell
# Configure
cmake -B build -G "Visual Studio 18 2026" -A x64

# Build
cmake --build build --config Release

# Install to the installation directory
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

## Custom 7-Zip Installation Path

The extension automatically detects 7-Zip installations by checking:
1. Registry paths: `HKEY_LOCAL_MACHINE\SOFTWARE\7-Zip-Zstandard` and `HKEY_LOCAL_MACHINE\SOFTWARE\7-Zip`
2. Common installation paths: `C:\Program Files\7-Zip` and `C:\Program Files (x86)\7-Zip`

7-Zip Zstandard is prioritized if both versions are installed. No manual configuration is needed unless 7-Zip is installed to a non-standard location.

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
- **COM Shell Extension**: Implements `IExplorerCommand` interface with dynamic submenu support
- **MSIX Sparse Package**: Allows non-store apps to register in Windows 11's new context menu
- **7-Zip SDK**: Direct integration with 7z.dll for compression/decompression and archive operations
- **Password Management**: Stores and recalls frequently used passwords for encrypted archives
- **Dynamic Detection**: Auto-detects 7-Zip installations via registry and file system
- **Self-signed certificate**: Automatically created and installed to the Trusted Root Certification Authorities store

## License

MIT License
