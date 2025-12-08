# 7-Zip Windows 11 Context Menu

将 7-Zip 添加到 Windows 11 **新右键菜单**（不需要点击"显示更多选项"），使用二级子菜单组织功能。

## 功能

右键点击文件或文件夹时，显示 **7-Zip** 菜单，包含以下子菜单：

### 压缩文件 (.7z, .zip, .rar 等)
- 提取到此处
- 提取到子文件夹

### 普通文件/文件夹
- 添加到 .7z
- 添加到 .zip

支持中英文自动切换。

## 系统要求

- Windows 11 (版本 21H2 或更高)
- [Visual Studio 2022](https://visualstudio.microsoft.com/) 或 [Build Tools](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022) (需要 C++ 桌面开发工作负载)
- [CMake](https://cmake.org/download/) 3.20+
- 7-Zip 安装在默认路径 `C:\Program Files\7-Zip`

## 构建与安装

### 1. 启用开发者模式

以 **管理员身份** 打开 PowerShell，运行：

```powershell
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /v AllowDevelopmentWithoutDevLicense /t REG_DWORD /d 1 /f
```

或者通过 GUI：设置 > 隐私和安全性 > 开发者选项 > 开启"开发人员模式"

### 2. 编译

打开 **"Developer Command Prompt for VS 2022"** 或 **PowerShell**：

```powershell
# 配置
cmake -B build -G "Visual Studio 17 2022" -A x64

# 编译
cmake --build build --config Release

# 安装到 install 目录
cmake --install build --config Release
```

### 3. 注册

以 **管理员身份** 打开 PowerShell：

```powershell
cd build/install
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

### 4. 重启资源管理器

```powershell
Stop-Process -Name explorer -Force; Start-Process explorer
```

或者注销并重新登录。

## 卸载

以 **管理员身份** 打开 PowerShell：

```powershell
cd build/install
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## 项目结构

```
7z/
├── CMakeLists.txt          # CMake 配置
├── README.md
├── src/
│   ├── ContextMenu.cpp     # IExplorerCommand COM 实现
│   ├── ContextMenu.h       # 头文件定义
│   └── ContextMenu.def     # DLL 导出
├── scripts/
│   ├── install.ps1         # 安装脚本
│   └── uninstall.ps1       # 卸载脚本
├── Package/                # 源资源文件
│   ├── AppxManifest.xml    # MSIX Sparse Package 清单
│   └── Assets/             # 图标目录
└── build/                  # 构建目录
    └── install/            # 安装输出
        ├── 7ZipContextMenu.dll
        ├── 7ZipContextMenu.pdb
        ├── AppxManifest.xml
        ├── Assets/
        ├── passwords.txt   # 密码存储文件
        ├── install.ps1
        └── uninstall.ps1
```

## 自定义 7-Zip 路径

如果 7-Zip 安装在其他位置，修改 `src/ContextMenu.cpp` 中的路径：

```cpp
static const wchar_t* g_7zPath = L"C:\\Program Files\\7-Zip";
static const wchar_t* g_7zDll = L"C:\\Program Files\\7-Zip\\7z.dll";
static const wchar_t* g_7zFM = L"C:\\Program Files\\7-Zip\\7zFM.exe";
```

然后重新编译并安装。

## 故障排除

### 包注册失败 (0x80073CFF)
- 确保已启用开发者模式（见步骤 1）
- 确保以管理员身份运行 install.ps1

### 菜单不显示
- 重启 Windows Explorer：`Stop-Process -Name explorer -Force; Start-Process explorer`
- 或者注销并重新登录

### 编译错误
- 确保已安装 Visual Studio 2022 的 C++ 桌面开发工作负载
- 确保 CMake 已添加到 PATH

### PowerShell 脚本无法运行
使用 `-ExecutionPolicy Bypass` 参数：
```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

## 技术原理

本项目使用：
- **COM Shell Extension**: 实现 `IExplorerCommand` 接口，支持子菜单
- **MSIX Sparse Package**: 允许非商店应用在 Windows 11 新右键菜单中注册
- **7-Zip SDK**: 直接调用 7z.dll 进行压缩/解压
- **自签名证书**: 自动创建并安装到受信任的根证书

## License

MIT License
