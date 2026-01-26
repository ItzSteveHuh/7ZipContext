#include <initguid.h>
#include "ContextMenu.h"
#include "SevenZipCore.h"
#include "ProgressDialog.h"
#include <pathcch.h>
#include <propvarutil.h>
#include <propkey.h>
#include <commctrl.h>
#include <algorithm>
#include <cstdio>
#include <thread>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "pathcch.lib")
#pragma comment(lib, "propsys.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "comctl32.lib")

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
    .extractHere = L"Extract Here",
    .extractTo = L"Extract to Subfolder",
    .addTo7z = L"Add to .7z",
    .addToZip = L"Add to .zip",
    .overwriteTitle = L"Confirm Overwrite",
    .overwriteMessage = L"Files already exist in the destination. Overwrite?",
    .passwordTitle = L"Extraction Complete",
    .passwordMessage = L"Archive extracted successfully.\n\nPassword used: %s\n\nYou can copy this password for future use.",
    .passwordPromptTitle = L"Password Required",
    .passwordPromptMessage = L"Enter password for encrypted archive:",
    .passwordWrongTitle = L"Wrong Password",
    .passwordWrongMessage = L"The password is incorrect. Try again?",
    .progressExtracting = L"Extracting...",
    .progressCompressing = L"Compressing...",
    .progressCancel = L"Cancel"
};

static const LocalizedStrings g_stringsCN = {
    .extractHere = L"\x63D0\x53D6\x5230\x6B64\x5904",           // 提取到此处
    .extractTo = L"\x63D0\x53D6\x5230\x5B50\x6587\x4EF6\x5939", // 提取到子文件夹
    .addTo7z = L"\x6DFB\x52A0\x5230 .7z",                       // 添加到 .7z
    .addToZip = L"\x6DFB\x52A0\x5230 .zip",                     // 添加到 .zip
    .overwriteTitle = L"\x786E\x8BA4\x8986\x76D6",              // 确认覆盖
    .overwriteMessage = L"\x76EE\x6807\x4F4D\x7F6E\x5DF2\x5B58\x5728\x6587\x4EF6\x3002\x662F\x5426\x8986\x76D6\xFF1F",
    .passwordTitle = L"\x89E3\x538B\x5B8C\x6210",               // 解压完成
    .passwordMessage = L"\x538B\x7F29\x5305\x89E3\x538B\x6210\x529F\x3002\n\n\x4F7F\x7528\x7684\x5BC6\x7801: %s\n\n\x60A8\x53EF\x4EE5\x590D\x5236\x6B64\x5BC6\x7801\x4EE5\x4FBF\x4E0B\x6B21\x4F7F\x7528\x3002",
    .passwordPromptTitle = L"\x9700\x8981\x5BC6\x7801",         // 需要密码
    .passwordPromptMessage = L"\x8BF7\x8F93\x5165\x52A0\x5BC6\x538B\x7F29\x5305\x7684\x5BC6\x7801:",
    .passwordWrongTitle = L"\x5BC6\x7801\x9519\x8BEF",          // 密码错误
    .passwordWrongMessage = L"\x5BC6\x7801\x4E0D\x6B63\x786E\x3002\x662F\x5426\x91CD\x8BD5\xFF1F",
    .progressExtracting = L"\x6B63\x5728\x89E3\x538B...",       // 正在解压...
    .progressCompressing = L"\x6B63\x5728\x538B\x7F29...",      // 正在压缩...
    .progressCancel = L"\x53D6\x6D88"                           // 取消
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
// Password Manager - stores passwords with usage count
//////////////////////////////////////////////////////////////////////////////

class CPasswordManager
{
public:
    static CPasswordManager& Instance() {
        static CPasswordManager instance;
        return instance;
    }

    void Load() {
        m_passwords.clear();
        std::wstring path = GetPasswordFilePath();

        FILE* f = _wfopen(path.c_str(), L"r, ccs=UTF-8");
        if (!f) return;

        wchar_t line[1024];
        while (fwscanf(f, L"%[^\n]\n", line) == 1) {
            wchar_t* tab = wcschr(line, L'\t');
            if (tab) {
                *tab = L'\0';
                int count = _wtoi(line);
                std::wstring pwd = tab + 1;
                m_passwords.push_back({ pwd, count });
            }
        }
        fclose(f);

        SortByCount();
    }

    void Save() {
        std::wstring path = GetPasswordFilePath();
        FILE* f = _wfopen(path.c_str(), L"w, ccs=UTF-8");
        if (!f) return;

        for (const auto& p : m_passwords) {
            fwprintf(f, L"%d\t%s\n", p.count, p.password.c_str());
        }
        fclose(f);
    }

    void IncrementCount(const std::wstring& password) {
        for (auto& p : m_passwords) {
            if (p.password == password) {
                p.count++;
                SortByCount();
                Save();
                return;
            }
        }
        // New password
        m_passwords.push_back({ password, 1 });
        SortByCount();
        Save();
    }

    const std::vector<std::wstring> GetPasswords() const {
        std::vector<std::wstring> result;
        for (const auto& p : m_passwords) {
            result.push_back(p.password);
        }
        return result;
    }

private:
    struct PasswordEntry {
        std::wstring password;
        int count;
    };
    std::vector<PasswordEntry> m_passwords;

    CPasswordManager() { Load(); }

    void SortByCount() {
        std::sort(m_passwords.begin(), m_passwords.end(),
            [](const PasswordEntry& a, const PasswordEntry& b) {
                return a.count > b.count;
            });
    }

    std::wstring GetPasswordFilePath() {
        wchar_t localAppData[MAX_PATH];
        if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, localAppData))) {
            std::wstring dir = std::wstring(localAppData) + L"\\7ZipContext";
            CreateDirectoryW(dir.c_str(), NULL);
            return dir + L"\\passwords.txt";
        }
        wchar_t dllPath[MAX_PATH];
        GetModuleFileNameW(g_hInst, dllPath, MAX_PATH);
        std::wstring path = dllPath;
        size_t pos = path.rfind(L'\\');
        if (pos != std::wstring::npos) {
            path = path.substr(0, pos);
        }
        return path + L"\\passwords.txt";
    }
};

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

// Password dialog data
static wchar_t g_passwordBuffer[256] = {};
static bool g_passwordDialogOK = false;

static INT_PTR CALLBACK PasswordDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
        case WM_INITDIALOG:
            SetFocus(GetDlgItem(hDlg, 101));
            return FALSE;
        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case IDOK:
                    GetDlgItemTextW(hDlg, 101, g_passwordBuffer, 256);
                    g_passwordDialogOK = true;
                    EndDialog(hDlg, IDOK);
                    return TRUE;
                case IDCANCEL:
                    g_passwordDialogOK = false;
                    EndDialog(hDlg, IDCANCEL);
                    return TRUE;
            }
            break;
        case WM_CLOSE:
            g_passwordDialogOK = false;
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
    }
    return FALSE;
}

static std::wstring ShowPasswordInputDialog()
{
    const auto& strings = GetLocalizedStrings();
    g_passwordBuffer[0] = L'\0';
    g_passwordDialogOK = false;

    #pragma pack(push, 4)
    struct {
        DLGTEMPLATE dlg;
        WORD menu, cls, title;
        wchar_t titleText[32];
        WORD itemType1;
        DLGITEMTEMPLATE item1;
        WORD item1Class1, item1Class2;
        wchar_t item1Text[64];
        WORD item1Extra;
        WORD itemType2;
        DLGITEMTEMPLATE item2;
        WORD item2Class1, item2Class2;
        WORD item2Text;
        WORD item2Extra;
        WORD itemType3;
        DLGITEMTEMPLATE item3;
        WORD item3Class1, item3Class2;
        wchar_t item3Text[8];
        WORD item3Extra;
        WORD itemType4;
        DLGITEMTEMPLATE item4;
        WORD item4Class1, item4Class2;
        wchar_t item4Text[8];
        WORD item4Extra;
    } dlgTemplate = {};
    #pragma pack(pop)

    dlgTemplate.dlg.style = DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU;
    dlgTemplate.dlg.dwExtendedStyle = 0;
    dlgTemplate.dlg.cdit = 4;
    dlgTemplate.dlg.x = 0;
    dlgTemplate.dlg.y = 0;
    dlgTemplate.dlg.cx = 200;
    dlgTemplate.dlg.cy = 80;
    dlgTemplate.menu = 0;
    dlgTemplate.cls = 0;
    dlgTemplate.title = 0;
    StringCchCopyW(dlgTemplate.titleText, 32, strings.passwordPromptTitle);

    dlgTemplate.item1.style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    dlgTemplate.item1.dwExtendedStyle = 0;
    dlgTemplate.item1.x = 10;
    dlgTemplate.item1.y = 10;
    dlgTemplate.item1.cx = 180;
    dlgTemplate.item1.cy = 12;
    dlgTemplate.item1.id = 100;
    dlgTemplate.item1Class1 = 0xFFFF;
    dlgTemplate.item1Class2 = 0x0082;
    StringCchCopyW(dlgTemplate.item1Text, 64, strings.passwordPromptMessage);
    dlgTemplate.item1Extra = 0;

    dlgTemplate.item2.style = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_PASSWORD | ES_AUTOHSCROLL;
    dlgTemplate.item2.dwExtendedStyle = 0;
    dlgTemplate.item2.x = 10;
    dlgTemplate.item2.y = 26;
    dlgTemplate.item2.cx = 180;
    dlgTemplate.item2.cy = 14;
    dlgTemplate.item2.id = 101;
    dlgTemplate.item2Class1 = 0xFFFF;
    dlgTemplate.item2Class2 = 0x0081;
    dlgTemplate.item2Text = 0;
    dlgTemplate.item2Extra = 0;

    dlgTemplate.item3.style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON;
    dlgTemplate.item3.dwExtendedStyle = 0;
    dlgTemplate.item3.x = 50;
    dlgTemplate.item3.y = 50;
    dlgTemplate.item3.cx = 45;
    dlgTemplate.item3.cy = 14;
    dlgTemplate.item3.id = IDOK;
    dlgTemplate.item3Class1 = 0xFFFF;
    dlgTemplate.item3Class2 = 0x0080;
    StringCchCopyW(dlgTemplate.item3Text, 8, L"OK");
    dlgTemplate.item3Extra = 0;

    dlgTemplate.item4.style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    dlgTemplate.item4.dwExtendedStyle = 0;
    dlgTemplate.item4.x = 105;
    dlgTemplate.item4.y = 50;
    dlgTemplate.item4.cx = 45;
    dlgTemplate.item4.cy = 14;
    dlgTemplate.item4.id = IDCANCEL;
    dlgTemplate.item4Class1 = 0xFFFF;
    dlgTemplate.item4Class2 = 0x0080;
    StringCchCopyW(dlgTemplate.item4Text, 8, IsChineseLocale() ? L"\x53D6\x6D88" : L"Cancel");
    dlgTemplate.item4Extra = 0;

    DialogBoxIndirectW(g_hInst, &dlgTemplate.dlg, NULL, PasswordDlgProc);

    if (g_passwordDialogOK && g_passwordBuffer[0] != L'\0') {
        return std::wstring(g_passwordBuffer);
    }
    return L"";
}

static void ShowPasswordResultDialog(const std::wstring& password)
{
    const auto& strings = GetLocalizedStrings();
    wchar_t message[1024];
    StringCchPrintfW(message, 1024, strings.passwordMessage, password.c_str());
    MessageBoxW(NULL, message, strings.passwordTitle, MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND);
}

static bool ConfirmOverwrite()
{
    const auto& strings = GetLocalizedStrings();
    int result = MessageBoxW(NULL, strings.overwriteMessage, strings.overwriteTitle,
                             MB_YESNO | MB_ICONQUESTION | MB_SETFOREGROUND);
    return (result == IDYES);
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
        case CommandType::ExtractHere:title = strings.extractHere; break;
        case CommandType::ExtractTo:  title = strings.extractTo; break;
        case CommandType::AddTo7z:    title = strings.addTo7z; break;
        case CommandType::AddToZip:   title = strings.addToZip; break;
    }

    return SHStrDupW(title, ppszName);
}

IFACEMETHODIMP CExplorerCommand::GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon)
{
    // Use a generic archive icon from shell32.dll
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

// Check if files exist in destination that would be overwritten
static bool CheckOverwriteNeededForArchive(const std::wstring& archivePath, const std::wstring& outDir)
{
    SevenZipCore& core = SevenZipCore::Instance();
    if (!core.OpenArchive(archivePath)) {
        return false;
    }

    bool needOverwrite = false;
    auto items = core.GetItems();
    for (const auto& item : items) {
        std::wstring fullPath = outDir;
        if (!fullPath.empty() && fullPath.back() != L'\\') {
            fullPath += L'\\';
        }
        fullPath += item.path;
        if (GetFileAttributesW(fullPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
            needOverwrite = true;
            break;
        }
    }

    core.CloseArchive();
    return needOverwrite;
}

bool CExplorerCommand::ExtractArchive(const std::wstring& archivePath, const std::wstring& outDir)
{
    SevenZipCore& core = SevenZipCore::Instance();
    CPasswordManager& pwdMgr = CPasswordManager::Instance();
    const auto& strings = GetLocalizedStrings();

    if (!core.OpenArchive(archivePath)) {
        return false;
    }

    std::wstring correctPassword;
    bool needsPassword = core.NeedsPassword();
    bool foundPassword = false;

    if (needsPassword) {
        // Try stored passwords
        std::vector<std::wstring> passwords = pwdMgr.GetPasswords();
        passwords.insert(passwords.begin(), L"");  // Try empty password first

        for (const auto& password : passwords) {
            if (core.TestPassword(password)) {
                correctPassword = password;
                foundPassword = true;
                break;
            }
        }

        // Prompt user if needed
        while (!foundPassword) {
            std::wstring userPassword = ShowPasswordInputDialog();
            if (userPassword.empty()) {
                core.CloseArchive();
                return false;  // User cancelled
            }

            if (core.TestPassword(userPassword)) {
                correctPassword = userPassword;
                foundPassword = true;
            } else {
                int result = MessageBoxW(NULL, strings.passwordWrongMessage, strings.passwordWrongTitle,
                                         MB_YESNO | MB_ICONWARNING | MB_SETFOREGROUND);
                if (result != IDYES) {
                    core.CloseArchive();
                    return false;
                }
            }
        }
    }

    // Show progress dialog
    CProgressDialog progressDlg;
    progressDlg.Show(strings.progressExtracting);

    bool cancelled = false;
    bool success = core.Extract(outDir, correctPassword,
        [&progressDlg, &cancelled](uint64_t completed, uint64_t total) -> bool {
            progressDlg.SetProgress(completed, total);
            // Handle pause
            while (progressDlg.IsPaused() && !progressDlg.IsCancelled()) {
                Sleep(100);
            }
            if (progressDlg.IsCancelled()) {
                cancelled = true;
                return false;
            }
            return true;
        });

    progressDlg.Close();
    core.CloseArchive();

    if (success && needsPassword && !correctPassword.empty()) {
        pwdMgr.IncrementCount(correctPassword);
        ShowPasswordResultDialog(correctPassword);
    }

    return success && !cancelled;
}

bool CExplorerCommand::CompressFiles(const std::vector<std::wstring>& srcPaths, const std::wstring& archivePath, const GUID& formatId)
{
    SevenZipCore& core = SevenZipCore::Instance();
    const auto& strings = GetLocalizedStrings();

    // Determine format from extension
    std::wstring ext = PathFindExtensionW(archivePath.c_str());
    std::wstring format = L"7z";
    if (_wcsicmp(ext.c_str(), L".zip") == 0) {
        format = L"zip";
    } else if (_wcsicmp(ext.c_str(), L".tar") == 0) {
        format = L"tar";
    }

    // Show progress dialog
    CProgressDialog progressDlg;
    progressDlg.Show(strings.progressCompressing);

    bool cancelled = false;
    bool success = core.Compress(srcPaths, archivePath, format,
        [&progressDlg, &cancelled](uint64_t completed, uint64_t total) -> bool {
            progressDlg.SetProgress(completed, total);
            // Handle pause
            while (progressDlg.IsPaused() && !progressDlg.IsCancelled()) {
                Sleep(100);
            }
            if (progressDlg.IsCancelled()) {
                cancelled = true;
                return false;
            }
            return true;
        });

    progressDlg.Close();

    return success && !cancelled;
}

IFACEMETHODIMP CExplorerCommand::Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc)
{
    GetSelectedItems(psiItemArray);
    if (m_selectedPaths.empty()) return E_FAIL;

    bool success = false;
    const std::wstring& firstPath = m_selectedPaths[0];

    // Dummy GUIDs for compatibility
    static const GUID CLSID_CFormat7z = { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x07, 0x00, 0x00 } };
    static const GUID CLSID_CFormatZip = { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x01, 0x00, 0x00 } };

    switch (m_type) {
        case CommandType::ExtractHere:
            {
                std::wstring outDir = GetParentDir(firstPath);
                if (CheckOverwriteNeededForArchive(firstPath, outDir)) {
                    if (!ConfirmOverwrite()) {
                        return S_OK;
                    }
                }
                success = ExtractArchive(firstPath, outDir);
            }
            break;

        case CommandType::ExtractTo:
            {
                std::wstring subDir = GetParentDir(firstPath) + L"\\" + GetFileNameWithoutExt(firstPath);
                if (GetFileAttributesW(subDir.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    if (CheckOverwriteNeededForArchive(firstPath, subDir)) {
                        if (!ConfirmOverwrite()) {
                            return S_OK;
                        }
                    }
                }
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
                success = CompressFiles(m_selectedPaths, archivePath, CLSID_CFormat7z);
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
                success = CompressFiles(m_selectedPaths, archivePath, CLSID_CFormatZip);
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
