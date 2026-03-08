#include <initguid.h>
#include "ContextMenu.h"
#include <winreg.h>

#pragma comment(lib, "shlwapi.lib")

// Global variables
HINSTANCE g_hInst = NULL;
LONG g_cDllRef = 0;

// Archive extensions
static const wchar_t* g_archiveExtensions[] = {
    L".7z", L".zip", L".rar", L".tar", L".gz", L".bz2",
    L".xz", L".iso", L".cab", L".arj", L".lzh", L".tgz",
    L".tbz2", L".wim"
};

// Localized strings
static const LocalizedStrings g_stringsEN = {
    .openArchive = L"Open Archive",
    .extractFiles = L"Extract Files...",
    .extractHere = L"Extract Here",
    .extractTo = L"Extract to Subfolder",
    .addTo7z = L"Add to .7z",
    .addToZip = L"Add to .zip"
};

static const LocalizedStrings g_stringsCN = {
    .openArchive = L"\x6253\x5F00\x538B\x7F29\x5305",         // 打开压缩包
    .extractFiles = L"\x89E3\x538B\x6587\x4EF6...",          // 解压文件...
    .extractHere = L"\x63D0\x53D6\x5230\x6B64\x5904",           // 提取到此处
    .extractTo = L"\x63D0\x53D6\x5230\x5B50\x6587\x4EF6\x5939", // 提取到子文件夹
    .addTo7z = L"\x6DFB\x52A0\x5230 .7z",                       // 添加到 .7z
    .addToZip = L"\x6DFB\x52A0\x5230 .zip"                      // 添加到 .zip
};

bool IsChineseLocale()
{
    LANGID langId = GetUserDefaultUILanguage();
    WORD primaryLang = PRIMARYLANGID(langId);
    return (primaryLang == LANG_CHINESE);
}

const LocalizedStrings& GetLocalizedStrings()
{
    return IsChineseLocale() ? g_stringsCN : g_stringsEN;
}

//////////////////////////////////////////////////////////////////////////////
// Helper functions
//////////////////////////////////////////////////////////////////////////////

static std::wstring GetFileName(const std::wstring& path)
{
    size_t pos = path.rfind(L'\\');
    return (pos != std::wstring::npos) ? path.substr(pos + 1) : path;
}

static std::wstring GetFileNameWithoutExt(const std::wstring& path)
{
    std::wstring name = GetFileName(path);
    size_t pos = name.rfind(L'.');
    return (pos != std::wstring::npos) ? name.substr(0, pos) : name;
}

static std::wstring GetParentDir(const std::wstring& path)
{
    size_t pos = path.rfind(L'\\');
    return (pos != std::wstring::npos) ? path.substr(0, pos) : path;
}

static std::wstring QuoteArg(const std::wstring& value)
{
    return L"\"" + value + L"\"";
}

static std::wstring TrimTrailingBackslash(std::wstring value)
{
    while (!value.empty() && (value.back() == L'\\' || value.back() == L'/')) {
        value.pop_back();
    }
    return value;
}

static std::wstring GetRegistryString(HKEY root, const wchar_t* subKey, const wchar_t* valueName)
{
    wchar_t buffer[MAX_PATH * 2] = {};
    DWORD size = sizeof(buffer);
    LONG rc = RegGetValueW(root, subKey, valueName, RRF_RT_REG_SZ, nullptr, buffer, &size);
    return (rc == ERROR_SUCCESS) ? std::wstring(buffer) : L"";
}

static bool PathExists(const std::wstring& path)
{
    return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;
}

static std::wstring Find7ZipExecutable()
{
    std::vector<std::wstring> baseDirs;

    std::wstring regPath = GetRegistryString(HKEY_LOCAL_MACHINE, L"SOFTWARE\\7-Zip", L"Path");
    if (!regPath.empty()) {
        baseDirs.push_back(TrimTrailingBackslash(regPath));
    }

    wchar_t pf[MAX_PATH] = {};
    DWORD pfLen = GetEnvironmentVariableW(L"ProgramFiles", pf, ARRAYSIZE(pf));
    if (pfLen > 0 && pfLen < ARRAYSIZE(pf)) {
        baseDirs.push_back(std::wstring(pf) + L"\\7-Zip");
    }

    wchar_t pfx86[MAX_PATH] = {};
    DWORD pfx86Len = GetEnvironmentVariableW(L"ProgramFiles(x86)", pfx86, ARRAYSIZE(pfx86));
    if (pfx86Len > 0 && pfx86Len < ARRAYSIZE(pfx86)) {
        baseDirs.push_back(std::wstring(pfx86) + L"\\7-Zip");
    }

    for (const auto& base : baseDirs) {
        std::wstring gui = base + L"\\7zG.exe";
        if (PathExists(gui)) {
            return gui;
        }
        std::wstring cli = base + L"\\7z.exe";
        if (PathExists(cli)) {
            return cli;
        }
    }

    return L"";
}

static std::wstring Find7ZipGuiExecutable()
{
    std::vector<std::wstring> baseDirs;

    std::wstring regPath = GetRegistryString(HKEY_LOCAL_MACHINE, L"SOFTWARE\\7-Zip", L"Path");
    if (!regPath.empty()) {
        baseDirs.push_back(TrimTrailingBackslash(regPath));
    }

    wchar_t pf[MAX_PATH] = {};
    DWORD pfLen = GetEnvironmentVariableW(L"ProgramFiles", pf, ARRAYSIZE(pf));
    if (pfLen > 0 && pfLen < ARRAYSIZE(pf)) {
        baseDirs.push_back(std::wstring(pf) + L"\\7-Zip");
    }

    wchar_t pfx86[MAX_PATH] = {};
    DWORD pfx86Len = GetEnvironmentVariableW(L"ProgramFiles(x86)", pfx86, ARRAYSIZE(pfx86));
    if (pfx86Len > 0 && pfx86Len < ARRAYSIZE(pfx86)) {
        baseDirs.push_back(std::wstring(pfx86) + L"\\7-Zip");
    }

    for (const auto& base : baseDirs) {
        std::wstring gui = base + L"\\7zG.exe";
        if (PathExists(gui)) {
            return gui;
        }
    }

    return L"";
}

static std::wstring Find7ZipFileManagerExecutable()
{
    std::vector<std::wstring> baseDirs;

    std::wstring regPath = GetRegistryString(HKEY_LOCAL_MACHINE, L"SOFTWARE\\7-Zip", L"Path");
    if (!regPath.empty()) {
        baseDirs.push_back(TrimTrailingBackslash(regPath));
    }

    wchar_t pf[MAX_PATH] = {};
    DWORD pfLen = GetEnvironmentVariableW(L"ProgramFiles", pf, ARRAYSIZE(pf));
    if (pfLen > 0 && pfLen < ARRAYSIZE(pf)) {
        baseDirs.push_back(std::wstring(pf) + L"\\7-Zip");
    }

    wchar_t pfx86[MAX_PATH] = {};
    DWORD pfx86Len = GetEnvironmentVariableW(L"ProgramFiles(x86)", pfx86, ARRAYSIZE(pfx86));
    if (pfx86Len > 0 && pfx86Len < ARRAYSIZE(pfx86)) {
        baseDirs.push_back(std::wstring(pfx86) + L"\\7-Zip");
    }

    for (const auto& base : baseDirs) {
        std::wstring fm = base + L"\\7zFM.exe";
        if (PathExists(fm)) {
            return fm;
        }
    }

    return L"";
}

static bool Run7Zip(const std::wstring& arguments)
{
    std::wstring exePath = Find7ZipExecutable();
    if (exePath.empty()) {
        MessageBoxW(NULL,
            IsChineseLocale() ? L"未找到 7-Zip。请先安装 7-Zip 到默认目录。" : L"7-Zip was not found. Please install 7-Zip in the default location.",
            L"7-Zip Context Menu",
            MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return false;
    }

    std::wstring cmdLine = QuoteArg(exePath) + L" " + arguments;
    std::vector<wchar_t> mutableCmd(cmdLine.begin(), cmdLine.end());
    mutableCmd.push_back(L'\0');

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi = {};

    if (!CreateProcessW(nullptr, mutableCmd.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &si, &pi)) {
        return false;
    }

    WaitForSingleObject(pi.hProcess, INFINITE);

    DWORD exitCode = 1;
    GetExitCodeProcess(pi.hProcess, &exitCode);

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);

    return exitCode == 0;
}

static bool OpenArchiveInFileManager(const std::wstring& archivePath)
{
    std::wstring exePath = Find7ZipFileManagerExecutable();
    if (exePath.empty()) {
        MessageBoxW(NULL,
            IsChineseLocale() ? L"未找到 7-Zip 文件管理器。请先安装完整的 7-Zip。" : L"7-Zip File Manager was not found. Please install full 7-Zip.",
            L"7-Zip Context Menu",
            MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return false;
    }

    std::wstring cmdLine = QuoteArg(exePath) + L" " + QuoteArg(archivePath);
    std::vector<wchar_t> mutableCmd(cmdLine.begin(), cmdLine.end());
    mutableCmd.push_back(L'\0');

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi = {};

    if (!CreateProcessW(nullptr, mutableCmd.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &si, &pi)) {
        return false;
    }

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return true;
}

static bool Run7ZipGui(const std::wstring& arguments, const std::wstring& workingDir = L"", bool waitForCompletion = true)
{
    std::wstring exePath = Find7ZipGuiExecutable();
    if (exePath.empty()) {
        MessageBoxW(NULL,
            IsChineseLocale() ? L"未找到 7-Zip 图形解压程序 (7zG.exe)。请安装完整的 7-Zip。" : L"7-Zip GUI extractor (7zG.exe) was not found. Please install full 7-Zip.",
            L"7-Zip Context Menu",
            MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return false;
    }

    std::wstring cmdLine = QuoteArg(exePath) + L" " + arguments;
    std::vector<wchar_t> mutableCmd(cmdLine.begin(), cmdLine.end());
    mutableCmd.push_back(L'\0');

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi = {};

    const wchar_t* lpCurrentDirectory = workingDir.empty() ? nullptr : workingDir.c_str();

    if (!CreateProcessW(nullptr, mutableCmd.data(), nullptr, nullptr, FALSE, 0, nullptr, lpCurrentDirectory, &si, &pi)) {
        return false;
    }

    if (waitForCompletion) {
        WaitForSingleObject(pi.hProcess, INFINITE);

        DWORD exitCode = 1;
        GetExitCodeProcess(pi.hProcess, &exitCode);

        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);

        return exitCode == 0;
    } else {
        // Fire-and-forget: let 7-Zip run in background without blocking
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
        return true;
    }
}

//////////////////////////////////////////////////////////////////////////////
// CExplorerCommand implementation
//////////////////////////////////////////////////////////////////////////////

CExplorerCommand::CExplorerCommand(CommandType type)
    : m_refCount(1), m_type(type), m_enumIndex(0), m_isArchive(false)
{
    InterlockedIncrement(&g_cDllRef);
    if (type == CommandType::Root) {
        InitSubCommands();
    }
}

CExplorerCommand::~CExplorerCommand()
{
    for (auto cmd : m_subCommands) {
        cmd->Release();
    }
    InterlockedDecrement(&g_cDllRef);
}

void CExplorerCommand::InitSubCommands()
{
}

bool CExplorerCommand::IsArchiveFile(const std::wstring& path)
{
    std::wstring ext = PathFindExtensionW(path.c_str());
    for (const auto& archiveExt : g_archiveExtensions) {
        if (_wcsicmp(ext.c_str(), archiveExt) == 0) {
            return true;
        }
    }
    return false;
}

void CExplorerCommand::GetSelectedItems(IShellItemArray* psiItemArray)
{
    m_selectedPaths.clear();
    m_isArchive = false;
    m_isDirectory = false;
    if (!psiItemArray) return;

    DWORD count = 0;
    psiItemArray->GetCount(&count);
    for (DWORD i = 0; i < count; i++) {
        IShellItem* psi = nullptr;
        if (SUCCEEDED(psiItemArray->GetItemAt(i, &psi))) {
            PWSTR pszPath = nullptr;
            if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath))) {
                m_selectedPaths.push_back(pszPath);
                CoTaskMemFree(pszPath);
            }
            psi->Release();
        }
    }
    if (!m_selectedPaths.empty()) {
        m_isArchive = IsArchiveFile(m_selectedPaths[0]);
        DWORD attrs = GetFileAttributesW(m_selectedPaths[0].c_str());
        m_isDirectory = (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY));
    }
}

IFACEMETHODIMP CExplorerCommand::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(CExplorerCommand, IExplorerCommand),
        QITABENT(CExplorerCommand, IEnumExplorerCommand),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) CExplorerCommand::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CExplorerCommand::Release()
{
    LONG ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) {
        delete this;
    }
    return ref;
}

IFACEMETHODIMP CExplorerCommand::GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName)
{
    const wchar_t* title = L"7-Zip";
    const auto& strings = GetLocalizedStrings();

    switch (m_type) {
        case CommandType::Root:       title = L"7-Zip"; break;
        case CommandType::OpenArchive:title = strings.openArchive; break;
        case CommandType::ExtractFiles:title = strings.extractFiles; break;
        case CommandType::ExtractHere:title = strings.extractHere; break;
        case CommandType::ExtractTo:
        {
            // Dynamically generate title with archive name
            GetSelectedItems(psiItemArray);
            if (!m_selectedPaths.empty()) {
                std::wstring archiveName = GetFileNameWithoutExt(m_selectedPaths[0]);
                std::wstring dynamicTitle = IsChineseLocale() ? 
                    (L"\u63D0\u53D6\u5230 \"" + archiveName + L"\\\"") :  // 提取到 "archiveName\"
                    (L"Extract to \"" + archiveName + L"\\\"");
                return SHStrDupW(dynamicTitle.c_str(), ppszName);
            }
            title = strings.extractTo;
            break;
        }
        case CommandType::AddTo7z:    title = strings.addTo7z; break;
        case CommandType::AddToZip:   title = strings.addToZip; break;
    }

    return SHStrDupW(title, ppszName);
}

IFACEMETHODIMP CExplorerCommand::GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon)
{
    std::wstring fmPath = Find7ZipFileManagerExecutable();
    if (!fmPath.empty()) {
        // Use the first icon resource from 7zFM.exe
        std::wstring iconRef = fmPath + L",0";
        return SHStrDupW(iconRef.c_str(), ppszIcon);
    }

    // Fallback when 7-Zip is not installed
    return SHStrDupW(L"shell32.dll,-171", ppszIcon);
}

IFACEMETHODIMP CExplorerCommand::GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip)
{
    *ppszInfotip = NULL;
    return E_NOTIMPL;
}

IFACEMETHODIMP CExplorerCommand::GetCanonicalName(GUID* pguidCommandName)
{
    if (m_type == CommandType::Root) {
        *pguidCommandName = CLSID_7ZipContextMenu;
        return S_OK;
    }
    ZeroMemory(pguidCommandName, sizeof(GUID));
    return E_NOTIMPL;
}

IFACEMETHODIMP CExplorerCommand::GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState)
{
    GetSelectedItems(psiItemArray);

    if (m_type == CommandType::Root) {
        *pCmdState = ECS_ENABLED;
        return S_OK;
    }

    switch (m_type) {
        case CommandType::OpenArchive:
        case CommandType::ExtractFiles:
            *pCmdState = m_isArchive ? ECS_ENABLED : ECS_HIDDEN;
            break;

        case CommandType::ExtractHere:
        case CommandType::ExtractTo:
            *pCmdState = (m_isArchive || !m_isDirectory) ? ECS_ENABLED : ECS_HIDDEN;
            break;

        case CommandType::AddTo7z:
        case CommandType::AddToZip:
            *pCmdState = m_isArchive ? ECS_HIDDEN : ECS_ENABLED;
            break;

        default:
            *pCmdState = ECS_ENABLED;
            break;
    }

    return S_OK;
}

bool CExplorerCommand::ExtractArchive(const std::wstring& archivePath, const std::wstring& outDir)
{
    std::wstring args = L"x " + QuoteArg(archivePath) + L" -o" + QuoteArg(outDir);
    return Run7Zip(args);
}

bool CExplorerCommand::CompressFiles(const std::vector<std::wstring>& srcPaths, const std::wstring& archivePath)
{
    std::wstring ext = PathFindExtensionW(archivePath.c_str());
    std::wstring format = L"7z";
    if (_wcsicmp(ext.c_str(), L".zip") == 0) {
        format = L"zip";
    } else if (_wcsicmp(ext.c_str(), L".tar") == 0) {
        format = L"tar";
    }

    std::wstring args = L"a -t" + format + L" " + QuoteArg(archivePath);
    for (const auto& srcPath : srcPaths) {
        args += L" " + QuoteArg(srcPath);
    }
    return Run7Zip(args);
}

IFACEMETHODIMP CExplorerCommand::Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc)
{
    GetSelectedItems(psiItemArray);
    if (m_selectedPaths.empty()) return E_FAIL;

    bool success = false;
    const std::wstring& firstPath = m_selectedPaths[0];

    switch (m_type) {
        case CommandType::OpenArchive:
            success = OpenArchiveInFileManager(firstPath);
            break;

        case CommandType::ExtractFiles:
            // Match 7-Zip shell behavior for "Extract files...":
            // x = extract, -o = default output path, -ad = show dialog,
            // -an/-ai = archive include switch.
            // Use parent + archive-name subfolder to mirror native 7-Zip field split.
            // Launch asynchronously (don't wait) to allow folder access during extraction.
            {
                std::wstring parentDir = GetParentDir(firstPath);
                std::wstring defaultOutDir = parentDir + L"\\" + GetFileNameWithoutExt(firstPath) + L"\\";
                success = Run7ZipGui(L"x -o" + QuoteArg(defaultOutDir) + L" -ad -an -ai!" + QuoteArg(firstPath), parentDir, false);
            }
            break;

        case CommandType::ExtractHere:
            {
                std::wstring outDir = GetParentDir(firstPath);
                success = ExtractArchive(firstPath, outDir);
            }
            break;

        case CommandType::ExtractTo:
            {
                std::wstring subDir = GetParentDir(firstPath) + L"\\" + GetFileNameWithoutExt(firstPath);
                CreateDirectoryW(subDir.c_str(), NULL);
                success = ExtractArchive(firstPath, subDir);
            }
            break;

        case CommandType::AddTo7z:
            {
                std::wstring archivePath;
                if (m_selectedPaths.size() == 1) {
                    archivePath = firstPath + L".7z";
                } else {
                    std::wstring parentDir = GetParentDir(firstPath);
                    std::wstring parentName = GetFileName(parentDir);
                    archivePath = parentDir + L"\\" + parentName + L".7z";
                }
                success = CompressFiles(m_selectedPaths, archivePath);
            }
            break;

        case CommandType::AddToZip:
            {
                std::wstring archivePath;
                if (m_selectedPaths.size() == 1) {
                    archivePath = firstPath + L".zip";
                } else {
                    std::wstring parentDir = GetParentDir(firstPath);
                    std::wstring parentName = GetFileName(parentDir);
                    archivePath = parentDir + L"\\" + parentName + L".zip";
                }
                success = CompressFiles(m_selectedPaths, archivePath);
            }
            break;

        default:
            return E_NOTIMPL;
    }

    return success ? S_OK : E_FAIL;
}

IFACEMETHODIMP CExplorerCommand::GetFlags(EXPCMDFLAGS* pFlags)
{
    if (m_type == CommandType::Root) {
        *pFlags = ECF_HASSUBCOMMANDS;
    } else {
        *pFlags = ECF_DEFAULT;
    }
    return S_OK;
}

IFACEMETHODIMP CExplorerCommand::EnumSubCommands(IEnumExplorerCommand** ppEnum)
{
    if (m_type != CommandType::Root) {
        *ppEnum = NULL;
        return E_NOTIMPL;
    }

    for (auto cmd : m_subCommands) {
        cmd->Release();
    }
    m_subCommands.clear();

    m_subCommands.push_back(new CExplorerCommand(CommandType::OpenArchive));
    m_subCommands.push_back(new CExplorerCommand(CommandType::ExtractFiles));
    m_subCommands.push_back(new CExplorerCommand(CommandType::ExtractHere));
    m_subCommands.push_back(new CExplorerCommand(CommandType::ExtractTo));
    m_subCommands.push_back(new CExplorerCommand(CommandType::AddTo7z));
    m_subCommands.push_back(new CExplorerCommand(CommandType::AddToZip));

    m_enumIndex = 0;
    this->AddRef();
    *ppEnum = this;
    return S_OK;
}

IFACEMETHODIMP CExplorerCommand::Next(ULONG celt, IExplorerCommand** pUICommand, ULONG* pceltFetched)
{
    ULONG fetched = 0;

    while (fetched < celt && m_enumIndex < m_subCommands.size()) {
        m_subCommands[m_enumIndex]->AddRef();
        pUICommand[fetched] = m_subCommands[m_enumIndex];
        m_enumIndex++;
        fetched++;
    }

    if (pceltFetched) {
        *pceltFetched = fetched;
    }

    return (fetched == celt) ? S_OK : S_FALSE;
}

IFACEMETHODIMP CExplorerCommand::Skip(ULONG celt)
{
    m_enumIndex += celt;
    return (m_enumIndex <= m_subCommands.size()) ? S_OK : S_FALSE;
}

IFACEMETHODIMP CExplorerCommand::Reset()
{
    m_enumIndex = 0;
    return S_OK;
}

IFACEMETHODIMP CExplorerCommand::Clone(IEnumExplorerCommand** ppenum)
{
    *ppenum = NULL;
    return E_NOTIMPL;
}

//////////////////////////////////////////////////////////////////////////////
// CClassFactory implementation
//////////////////////////////////////////////////////////////////////////////

CClassFactory::CClassFactory() : m_refCount(1)
{
    InterlockedIncrement(&g_cDllRef);
}

CClassFactory::~CClassFactory()
{
    InterlockedDecrement(&g_cDllRef);
}

IFACEMETHODIMP CClassFactory::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(CClassFactory, IClassFactory),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) CClassFactory::AddRef()
{
    return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) CClassFactory::Release()
{
    LONG ref = InterlockedDecrement(&m_refCount);
    if (ref == 0) {
        delete this;
    }
    return ref;
}

IFACEMETHODIMP CClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
    if (pUnkOuter != NULL) {
        return CLASS_E_NOAGGREGATION;
    }

    CExplorerCommand* pCmd = new (std::nothrow) CExplorerCommand(CommandType::Root);
    if (pCmd == NULL) {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pCmd->QueryInterface(riid, ppv);
    pCmd->Release();
    return hr;
}

IFACEMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
    if (fLock) {
        InterlockedIncrement(&g_cDllRef);
    } else {
        InterlockedDecrement(&g_cDllRef);
    }
    return S_OK;
}

//////////////////////////////////////////////////////////////////////////////
// DLL exports
//////////////////////////////////////////////////////////////////////////////

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (!IsEqualCLSID(rclsid, CLSID_7ZipContextMenu)) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    CClassFactory* pFactory = new (std::nothrow) CClassFactory();
    if (pFactory == NULL) {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pFactory->QueryInterface(riid, ppv);
    pFactory->Release();
    return hr;
}

STDAPI DllCanUnloadNow()
{
    return (g_cDllRef > 0) ? S_FALSE : S_OK;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call) {
        case DLL_PROCESS_ATTACH:
            g_hInst = hModule;
            DisableThreadLibraryCalls(hModule);
            break;
        case DLL_PROCESS_DETACH:
            break;
    }
    return TRUE;
}
