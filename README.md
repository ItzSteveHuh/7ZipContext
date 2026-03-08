# 7-Zip Windows 11 Context Menu

Add 7-Zip to the Windows 11 **new context menu** (without clicking "Show more options"), with features organized in a second-level submenu.

## Features

When you right-click a file or folder, a **7-Zip** menu appears:

- **Archive files** (.7z, .zip, .rar, etc.): Extract Here, Extract to Subfolder
- **Regular files/folders**: Add to .7z, Add to .zip

Supports automatic switching between Chinese and English.

## Installation

### Option 1: Download the installer (recommended)

Download the MSI installer from [Releases](https://github.com/YOUR_USERNAME/7ZipContext/releases) and double-click to install.

If double-click installation fails, install from the command line (open PowerShell as **Administrator**):

```powershell
msiexec /i "7ZipContext-x64.msi" /l*v "$env:USERPROFILE\Desktop\msi.log"
```

### Option 2: Build from source

**System requirements**:
- Windows 11 (21H2+)
- [Visual Studio 2022](https://visualstudio.microsoft.com/) or Build Tools (Desktop development with C++)
- [CMake](https://cmake.org/download/) 3.20+
- 7-Zip installed at `C:\Program Files\7-Zip`

**Steps**:

1. Enable Developer Mode (Administrator PowerShell):
   ```powershell
   reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /v AllowDevelopmentWithoutDevLicense /t REG_DWORD /d 1 /f
   ```

2. Build and install:
   ```powershell
   cmake -B build -G "Visual Studio 17 2022" -A x64
   cmake --build build --config Release
   cmake --install build --config Release
   ```

3. Register (Administrator PowerShell):
   ```powershell
   cd build/install
   powershell -ExecutionPolicy Bypass -File .\install.ps1
   ```

4. Restart Explorer:
   ```powershell
   Stop-Process -Name explorer -Force; Start-Process explorer
   ```

## Uninstall

Run as **Administrator**:
```powershell
cd build/install
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## Troubleshooting

| Problem | Solution |
|------|---------|
| Package registration failed (0x80073CFF) | Enable Developer Mode and run as Administrator |
| Menu does not appear | Restart Explorer or sign out and sign in again |
| Custom 7-Zip path | Modify the path in `src/ContextMenu.cpp` and rebuild |

## License

MIT License
