# 7-Zip Windows 11 Context Menu

将 7-Zip 添加到 Windows 11 **新右键菜单**（不需要点击"显示更多选项"），使用二级子菜单组织功能。

## 功能

右键点击文件或文件夹时，显示 **7-Zip** 菜单：

- **压缩文件** (.7z, .zip, .rar 等): 提取到此处、提取到子文件夹
- **普通文件/文件夹**: 添加到 .7z、添加到 .zip

支持中英文自动切换。

## 安装

### 方式一：下载安装包（推荐）

从 [Releases](https://github.com/YOUR_USERNAME/7ZipContext/releases) 下载 MSI 安装包，双击安装。

如果双击安装失败，使用命令行安装（以**管理员身份**打开 PowerShell）：

```powershell
msiexec /i "7ZipContext-x64.msi" /l*v "$env:USERPROFILE\Desktop\msi.log"
```

### 方式二：从源码构建

**系统要求**：
- Windows 11 (21H2+)
- [Visual Studio 2022](https://visualstudio.microsoft.com/) 或 Build Tools (C++ 桌面开发)
- [CMake](https://cmake.org/download/) 3.20+
- 7-Zip 安装在 `C:\Program Files\7-Zip`

**步骤**：

1. 启用开发者模式（管理员 PowerShell）：
   ```powershell
   reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /v AllowDevelopmentWithoutDevLicense /t REG_DWORD /d 1 /f
   ```

2. 编译安装：
   ```powershell
   cmake -B build -G "Visual Studio 17 2022" -A x64
   cmake --build build --config Release
   cmake --install build --config Release
   ```

3. 注册（管理员 PowerShell）：
   ```powershell
   cd build/install
   powershell -ExecutionPolicy Bypass -File .\install.ps1
   ```

4. 重启资源管理器：
   ```powershell
   Stop-Process -Name explorer -Force; Start-Process explorer
   ```

## 卸载

**管理员身份**运行：
```powershell
cd build/install
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| 包注册失败 (0x80073CFF) | 启用开发者模式，以管理员身份运行 |
| 菜单不显示 | 重启资源管理器或注销重登 |
| 自定义 7-Zip 路径 | 修改 `src/ContextMenu.cpp` 中的路径后重新编译 |

## License

MIT License
