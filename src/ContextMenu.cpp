#include <initguid.h>
#include "ContextMenu.h"
#include <pathcch.h>
#include <propvarutil.h>
#include <propkey.h>
#include <algorithm>
#include <cstdio>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "pathcch.lib")
#pragma comment(lib, "propsys.lib")
#pragma comment(lib, "oleaut32.lib")

// 7-Zip Archive GUIDs
DEFINE_GUID(CLSID_CFormat7z, 0x23170F69, 0x40C1, 0x278A, 0x10, 0x00, 0x00, 0x01, 0x10, 0x07, 0x00, 0x00);
DEFINE_GUID(CLSID_CFormatZip, 0x23170F69, 0x40C1, 0x278A, 0x10, 0x00, 0x00, 0x01, 0x10, 0x01, 0x00, 0x00);

// 7-Zip Interface IDs
DEFINE_GUID(IID_IInArchive, 0x23170F69, 0x40C1, 0x278A, 0x00, 0x00, 0x00, 0x06, 0x00, 0x60, 0x00, 0x00);
DEFINE_GUID(IID_IOutArchive, 0x23170F69, 0x40C1, 0x278A, 0x00, 0x00, 0x00, 0x06, 0x00, 0xA0, 0x00, 0x00);

// Global variables
HINSTANCE g_hInst = NULL;
LONG g_cDllRef = 0;

// 7-Zip paths - will be detected dynamically
static std::wstring g_7zExePath;  // 7z.exe path
static std::wstring g_7zDllPath;  // 7z.dll path
static std::wstring g_7zFMPath;   // 7zFM.exe path (for icon)

// Detect 7-Zip installation path
static std::wstring Detect7ZipPath()
{
    // Try registry first
    HKEY hKey;
    const wchar_t* regPaths[] = {
        L"SOFTWARE\\7-Zip",
        L"SOFTWARE\\WOW6432Node\\7-Zip"
    };

    for (const auto& regPath : regPaths) {
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, regPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
            wchar_t path[MAX_PATH];
            DWORD size = sizeof(path);
            if (RegQueryValueExW(hKey, L"Path", NULL, NULL, (LPBYTE)path, &size) == ERROR_SUCCESS) {
                RegCloseKey(hKey);
                return path;
            }
            RegCloseKey(hKey);
        }
    }

    // Try common paths
    const wchar_t* commonPaths[] = {
        L"C:\\Program Files\\7-Zip",
        L"C:\\Program Files (x86)\\7-Zip"
    };

    for (const auto& path : commonPaths) {
        std::wstring exePath = std::wstring(path) + L"\\7z.exe";
        if (GetFileAttributesW(exePath.c_str()) != INVALID_FILE_ATTRIBUTES) {
            return path;
        }
    }

    return L"";
}

// Initialize 7-Zip paths
static void Init7ZipPaths()
{
    static bool initialized = false;
    if (initialized) return;

    std::wstring basePath = Detect7ZipPath();
    if (!basePath.empty()) {
        // Remove trailing backslash if present
        if (basePath.back() == L'\\') {
            basePath.pop_back();
        }
        g_7zExePath = basePath + L"\\7z.exe";
        g_7zDllPath = basePath + L"\\7z.dll";
        g_7zFMPath = basePath + L"\\7zFM.exe";
    }
    initialized = true;
}

// Archive extensions with their format GUIDs
struct ArchiveFormat {
    const wchar_t* ext;
    const GUID* formatId;
};

static const ArchiveFormat g_archiveFormats[] = {
    { L".7z", &CLSID_CFormat7z },
    { L".zip", &CLSID_CFormatZip },
    { L".rar", nullptr },  // RAR uses auto-detection
    { L".tar", nullptr },
    { L".gz", nullptr },
    { L".bz2", nullptr },
    { L".xz", nullptr },
    { L".iso", nullptr },
    { L".cab", nullptr },
    { L".arj", nullptr },
    { L".lzh", nullptr },
    { L".tgz", nullptr },
    { L".tbz2", nullptr },
    { L".wim", nullptr }
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
    .passwordMessage = L"Archive extracted successfully.\n\nPassword used: %s\n\nYou can copy this password for future use."
};

static const LocalizedStrings g_stringsCN = {
    .extractHere = L"\x63D0\x53D6\x5230\x6B64\x5904",           // 提取到此处
    .extractTo = L"\x63D0\x53D6\x5230\x5B50\x6587\x4EF6\x5939", // 提取到子文件夹
    .addTo7z = L"\x6DFB\x52A0\x5230 .7z",                       // 添加到 .7z
    .addToZip = L"\x6DFB\x52A0\x5230 .zip",                     // 添加到 .zip
    .overwriteTitle = L"\x786E\x8BA4\x8986\x76D6",              // 确认覆盖
    .overwriteMessage = L"\x76EE\x6807\x4F4D\x7F6E\x5DF2\x5B58\x5728\x6587\x4EF6\x3002\x662F\x5426\x8986\x76D6\xFF1F", // 目标位置已存在文件。是否覆盖？
    .passwordTitle = L"\x89E3\x538B\x5B8C\x6210",               // 解压完成
    .passwordMessage = L"\x538B\x7F29\x5305\x89E3\x538B\x6210\x529F\x3002\n\n\x4F7F\x7528\x7684\x5BC6\x7801: %s\n\n\x60A8\x53EF\x4EE5\x590D\x5236\x6B64\x5BC6\x7801\x4EE5\x4FBF\x4E0B\x6B21\x4F7F\x7528\x3002" // 压缩包解压成功。\n\n使用的密码: %s\n\n您可以复制此密码以便下次使用。
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
// 7-Zip SDK Integration
//////////////////////////////////////////////////////////////////////////////

// 7-Zip SDK types
typedef UINT32 (WINAPI *Func_CreateObject)(const GUID *clsid, const GUID *iid, void **outObject);
typedef UINT32 (WINAPI *Func_GetNumberOfFormats)(UINT32 *numFormats);
typedef UINT32 (WINAPI *Func_GetHandlerProperty2)(UINT32 index, PROPID propID, PROPVARIANT *value);

// Property IDs
const PROPID kpidPath = 3;
const PROPID kpidName = 4;
const PROPID kpidIsDir = 6;
const PROPID kpidSize = 7;
const PROPID kpidAttrib = 9;
const PROPID kpidMTime = 14;

// 7-Zip interfaces
MIDL_INTERFACE("23170F69-40C1-278A-0000-000600100000")
ISequentialInStream : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE Read(void *data, UINT32 size, UINT32 *processedSize) = 0;
};

MIDL_INTERFACE("23170F69-40C1-278A-0000-000600200000")
ISequentialOutStream : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE Write(const void *data, UINT32 size, UINT32 *processedSize) = 0;
};

MIDL_INTERFACE("23170F69-40C1-278A-0000-000600030000")
IInStream : public ISequentialInStream
{
public:
    virtual HRESULT STDMETHODCALLTYPE Seek(INT64 offset, UINT32 seekOrigin, UINT64 *newPosition) = 0;
};

MIDL_INTERFACE("23170F69-40C1-278A-0000-000600040000")
IOutStream : public ISequentialOutStream
{
public:
    virtual HRESULT STDMETHODCALLTYPE Seek(INT64 offset, UINT32 seekOrigin, UINT64 *newPosition) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetSize(UINT64 newSize) = 0;
};

MIDL_INTERFACE("23170F69-40C1-278A-0000-000600600000")
IInArchive : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE Open(IInStream *stream, const UINT64 *maxCheckStartPosition, IUnknown *openCallback) = 0;
    virtual HRESULT STDMETHODCALLTYPE Close() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetNumberOfItems(UINT32 *numItems) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetProperty(UINT32 index, PROPID propID, PROPVARIANT *value) = 0;
    virtual HRESULT STDMETHODCALLTYPE Extract(const UINT32 *indices, UINT32 numItems, INT32 testMode, IUnknown *extractCallback) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetArchiveProperty(PROPID propID, PROPVARIANT *value) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetNumberOfProperties(UINT32 *numProps) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPropertyInfo(UINT32 index, BSTR *name, PROPID *propID, VARTYPE *varType) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetNumberOfArchiveProperties(UINT32 *numProps) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetArchivePropertyInfo(UINT32 index, BSTR *name, PROPID *propID, VARTYPE *varType) = 0;
};

MIDL_INTERFACE("23170F69-40C1-278A-0000-000600A00000")
IOutArchive : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE UpdateItems(ISequentialOutStream *outStream, UINT32 numItems, IUnknown *updateCallback) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetFileTimeType(UINT32 *type) = 0;
};

MIDL_INTERFACE("23170F69-40C1-278A-0000-000600200000")
IArchiveExtractCallback : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE SetTotal(UINT64 total) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetCompleted(const UINT64 *completeValue) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetStream(UINT32 index, ISequentialOutStream **outStream, INT32 askExtractMode) = 0;
    virtual HRESULT STDMETHODCALLTYPE PrepareOperation(INT32 askExtractMode) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetOperationResult(INT32 opRes) = 0;
};

MIDL_INTERFACE("23170F69-40C1-278A-0000-000600800000")
IArchiveUpdateCallback : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE SetTotal(UINT64 total) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetCompleted(const UINT64 *completeValue) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetUpdateItemInfo(UINT32 index, INT32 *newData, INT32 *newProps, UINT32 *indexInArchive) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetProperty(UINT32 index, PROPID propID, PROPVARIANT *value) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetStream(UINT32 index, ISequentialInStream **inStream) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetOperationResult(INT32 operationResult) = 0;
};

// Password interface for encrypted archives
MIDL_INTERFACE("23170F69-40C1-278A-0000-000500100000")
ICryptoGetTextPassword : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE CryptoGetTextPassword(BSTR *password) = 0;
};

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

// File stream implementation for IInStream
class CFileInStream : public IInStream
{
public:
    CFileInStream() : m_refCount(1), m_hFile(INVALID_HANDLE_VALUE) {}
    ~CFileInStream() { Close(); }

    bool Open(const wchar_t* path) {
        m_hFile = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
        return m_hFile != INVALID_HANDLE_VALUE;
    }
    void Close() {
        if (m_hFile != INVALID_HANDLE_VALUE) {
            CloseHandle(m_hFile);
            m_hFile = INVALID_HANDLE_VALUE;
        }
    }

    STDMETHOD(QueryInterface)(REFIID riid, void** ppv) {
        if (riid == IID_IUnknown || riid == __uuidof(ISequentialInStream) || riid == __uuidof(IInStream)) {
            *ppv = this;
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    STDMETHOD_(ULONG, AddRef)() { return InterlockedIncrement(&m_refCount); }
    STDMETHOD_(ULONG, Release)() {
        LONG ref = InterlockedDecrement(&m_refCount);
        if (ref == 0) delete this;
        return ref;
    }

    STDMETHOD(Read)(void* data, UINT32 size, UINT32* processedSize) {
        DWORD read = 0;
        BOOL result = ReadFile(m_hFile, data, size, &read, NULL);
        if (processedSize) *processedSize = read;
        return result ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }
    STDMETHOD(Seek)(INT64 offset, UINT32 seekOrigin, UINT64* newPosition) {
        LARGE_INTEGER li;
        li.QuadPart = offset;
        LARGE_INTEGER newPos;
        if (!SetFilePointerEx(m_hFile, li, &newPos, seekOrigin))
            return HRESULT_FROM_WIN32(GetLastError());
        if (newPosition) *newPosition = newPos.QuadPart;
        return S_OK;
    }

private:
    LONG m_refCount;
    HANDLE m_hFile;
};

// File stream implementation for IOutStream
class CFileOutStream : public IOutStream
{
public:
    CFileOutStream() : m_refCount(1), m_hFile(INVALID_HANDLE_VALUE) {}
    ~CFileOutStream() { Close(); }

    bool Create(const wchar_t* path) {
        m_hFile = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        return m_hFile != INVALID_HANDLE_VALUE;
    }
    void Close() {
        if (m_hFile != INVALID_HANDLE_VALUE) {
            CloseHandle(m_hFile);
            m_hFile = INVALID_HANDLE_VALUE;
        }
    }

    STDMETHOD(QueryInterface)(REFIID riid, void** ppv) {
        if (riid == IID_IUnknown || riid == __uuidof(ISequentialOutStream) || riid == __uuidof(IOutStream)) {
            *ppv = this;
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    STDMETHOD_(ULONG, AddRef)() { return InterlockedIncrement(&m_refCount); }
    STDMETHOD_(ULONG, Release)() {
        LONG ref = InterlockedDecrement(&m_refCount);
        if (ref == 0) delete this;
        return ref;
    }

    STDMETHOD(Write)(const void* data, UINT32 size, UINT32* processedSize) {
        DWORD written = 0;
        BOOL result = WriteFile(m_hFile, data, size, &written, NULL);
        if (processedSize) *processedSize = written;
        return result ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }
    STDMETHOD(Seek)(INT64 offset, UINT32 seekOrigin, UINT64* newPosition) {
        LARGE_INTEGER li;
        li.QuadPart = offset;
        LARGE_INTEGER newPos;
        if (!SetFilePointerEx(m_hFile, li, &newPos, seekOrigin))
            return HRESULT_FROM_WIN32(GetLastError());
        if (newPosition) *newPosition = newPos.QuadPart;
        return S_OK;
    }
    STDMETHOD(SetSize)(UINT64 newSize) {
        LARGE_INTEGER li;
        li.QuadPart = newSize;
        if (!SetFilePointerEx(m_hFile, li, NULL, FILE_BEGIN))
            return HRESULT_FROM_WIN32(GetLastError());
        if (!SetEndOfFile(m_hFile))
            return HRESULT_FROM_WIN32(GetLastError());
        return S_OK;
    }

private:
    LONG m_refCount;
    HANDLE m_hFile;
};

// Extract callback implementation with password support
class CExtractCallback : public IArchiveExtractCallback, public ICryptoGetTextPassword
{
public:
    CExtractCallback(const std::wstring& outDir, IInArchive* archive, const std::wstring& password = L"")
        : m_refCount(1), m_outDir(outDir), m_archive(archive), m_currentIndex(0),
          m_password(password), m_passwordWasUsed(false) {}

    bool WasPasswordUsed() const { return m_passwordWasUsed; }

    STDMETHOD(QueryInterface)(REFIID riid, void** ppv) {
        if (riid == IID_IUnknown || riid == __uuidof(IArchiveExtractCallback)) {
            *ppv = static_cast<IArchiveExtractCallback*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(ICryptoGetTextPassword)) {
            *ppv = static_cast<ICryptoGetTextPassword*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    STDMETHOD_(ULONG, AddRef)() { return InterlockedIncrement(&m_refCount); }
    STDMETHOD_(ULONG, Release)() {
        LONG ref = InterlockedDecrement(&m_refCount);
        if (ref == 0) delete this;
        return ref;
    }

    STDMETHOD(SetTotal)(UINT64 total) { return S_OK; }
    STDMETHOD(SetCompleted)(const UINT64* completeValue) { return S_OK; }

    STDMETHOD(GetStream)(UINT32 index, ISequentialOutStream** outStream, INT32 askExtractMode) {
        m_currentIndex = index;
        *outStream = NULL;
        if (askExtractMode != 0) return S_OK;  // Not extracting

        PROPVARIANT prop;
        PropVariantInit(&prop);
        m_archive->GetProperty(index, kpidPath, &prop);
        std::wstring itemPath = (prop.vt == VT_BSTR) ? prop.bstrVal : L"";
        PropVariantClear(&prop);

        PropVariantInit(&prop);
        m_archive->GetProperty(index, kpidIsDir, &prop);
        bool isDir = (prop.vt == VT_BOOL && prop.boolVal != VARIANT_FALSE);
        PropVariantClear(&prop);

        std::wstring fullPath = m_outDir + L"\\" + itemPath;

        if (isDir) {
            CreateDirectoryRecursive(fullPath);
            return S_OK;
        }

        // Create parent directory
        size_t pos = fullPath.rfind(L'\\');
        if (pos != std::wstring::npos) {
            CreateDirectoryRecursive(fullPath.substr(0, pos));
        }

        CFileOutStream* stream = new CFileOutStream();
        if (!stream->Create(fullPath.c_str())) {
            delete stream;
            return HRESULT_FROM_WIN32(GetLastError());
        }
        *outStream = stream;
        return S_OK;
    }

    STDMETHOD(PrepareOperation)(INT32 askExtractMode) { return S_OK; }
    STDMETHOD(SetOperationResult)(INT32 opRes) { return S_OK; }

    // ICryptoGetTextPassword
    STDMETHOD(CryptoGetTextPassword)(BSTR* password) {
        m_passwordWasUsed = true;
        if (m_password.empty()) {
            *password = NULL;
            return E_ABORT;  // No password available
        }
        *password = SysAllocString(m_password.c_str());
        return S_OK;
    }

private:
    LONG m_refCount;
    std::wstring m_outDir;
    IInArchive* m_archive;
    UINT32 m_currentIndex;
    std::wstring m_password;
    bool m_passwordWasUsed;

    void CreateDirectoryRecursive(const std::wstring& path) {
        size_t pos = 0;
        while ((pos = path.find(L'\\', pos + 1)) != std::wstring::npos) {
            CreateDirectoryW(path.substr(0, pos).c_str(), NULL);
        }
        CreateDirectoryW(path.c_str(), NULL);
    }
};

// Update callback implementation for compression
class CUpdateCallback : public IArchiveUpdateCallback
{
public:
    CUpdateCallback(const std::wstring& srcPath, const std::wstring& itemName)
        : m_refCount(1), m_srcPath(srcPath), m_itemName(itemName), m_isDir(false) {
        DWORD attrs = GetFileAttributesW(srcPath.c_str());
        m_isDir = (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY));
        if (m_isDir) {
            EnumerateFiles(srcPath, L"");
        }
    }

    STDMETHOD(QueryInterface)(REFIID riid, void** ppv) {
        if (riid == IID_IUnknown || riid == __uuidof(IArchiveUpdateCallback)) {
            *ppv = this;
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    STDMETHOD_(ULONG, AddRef)() { return InterlockedIncrement(&m_refCount); }
    STDMETHOD_(ULONG, Release)() {
        LONG ref = InterlockedDecrement(&m_refCount);
        if (ref == 0) delete this;
        return ref;
    }

    STDMETHOD(SetTotal)(UINT64 total) { return S_OK; }
    STDMETHOD(SetCompleted)(const UINT64* completeValue) { return S_OK; }

    STDMETHOD(GetUpdateItemInfo)(UINT32 index, INT32* newData, INT32* newProps, UINT32* indexInArchive) {
        if (newData) *newData = 1;  // New data
        if (newProps) *newProps = 1;  // New properties
        if (indexInArchive) *indexInArchive = (UINT32)-1;  // Not in archive
        return S_OK;
    }

    STDMETHOD(GetProperty)(UINT32 index, PROPID propID, PROPVARIANT* value) {
        PropVariantInit(value);

        std::wstring itemPath;
        bool itemIsDir = false;
        std::wstring fullPath;

        if (m_isDir) {
            if (index >= m_files.size()) return E_INVALIDARG;
            itemPath = m_files[index].relativePath;
            itemIsDir = m_files[index].isDir;
            fullPath = m_files[index].fullPath;
        } else {
            itemPath = m_itemName;
            itemIsDir = false;
            fullPath = m_srcPath;
        }

        WIN32_FILE_ATTRIBUTE_DATA fad = {};
        GetFileAttributesExW(fullPath.c_str(), GetFileExInfoStandard, &fad);

        switch (propID) {
            case kpidPath:
                value->vt = VT_BSTR;
                value->bstrVal = SysAllocString(itemPath.c_str());
                break;
            case kpidIsDir:
                value->vt = VT_BOOL;
                value->boolVal = itemIsDir ? VARIANT_TRUE : VARIANT_FALSE;
                break;
            case kpidSize:
                if (!itemIsDir) {
                    value->vt = VT_UI8;
                    value->uhVal.LowPart = fad.nFileSizeLow;
                    value->uhVal.HighPart = fad.nFileSizeHigh;
                }
                break;
            case kpidAttrib:
                value->vt = VT_UI4;
                value->ulVal = fad.dwFileAttributes;
                break;
            case kpidMTime:
                value->vt = VT_FILETIME;
                value->filetime = fad.ftLastWriteTime;
                break;
        }
        return S_OK;
    }

    STDMETHOD(GetStream)(UINT32 index, ISequentialInStream** inStream) {
        *inStream = NULL;

        std::wstring fullPath;
        bool itemIsDir;

        if (m_isDir) {
            if (index >= m_files.size()) return E_INVALIDARG;
            fullPath = m_files[index].fullPath;
            itemIsDir = m_files[index].isDir;
        } else {
            fullPath = m_srcPath;
            itemIsDir = false;
        }

        if (itemIsDir) return S_OK;

        CFileInStream* stream = new CFileInStream();
        if (!stream->Open(fullPath.c_str())) {
            delete stream;
            return HRESULT_FROM_WIN32(GetLastError());
        }
        *inStream = stream;
        return S_OK;
    }

    STDMETHOD(SetOperationResult)(INT32 operationResult) { return S_OK; }

    UINT32 GetItemCount() const { return m_isDir ? (UINT32)m_files.size() : 1; }

private:
    LONG m_refCount;
    std::wstring m_srcPath;
    std::wstring m_itemName;
    bool m_isDir;

    struct FileItem {
        std::wstring fullPath;
        std::wstring relativePath;
        bool isDir;
    };
    std::vector<FileItem> m_files;

    void EnumerateFiles(const std::wstring& basePath, const std::wstring& relPath) {
        std::wstring searchPath = basePath + L"\\*";
        WIN32_FIND_DATAW fd;
        HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
        if (hFind == INVALID_HANDLE_VALUE) return;

        do {
            if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0)
                continue;

            std::wstring relItemPath = relPath.empty() ? fd.cFileName : relPath + L"\\" + fd.cFileName;
            std::wstring fullItemPath = basePath + L"\\" + fd.cFileName;

            FileItem item;
            item.fullPath = fullItemPath;
            item.relativePath = m_itemName + L"\\" + relItemPath;
            item.isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            m_files.push_back(item);

            if (item.isDir) {
                EnumerateFiles(fullItemPath, relItemPath);
            }
        } while (FindNextFileW(hFind, &fd));
        FindClose(hFind);
    }
};

// 7-Zip library wrapper
class C7ZipLib
{
public:
    static C7ZipLib& Instance() {
        static C7ZipLib instance;
        return instance;
    }

    bool IsLoaded() const { return m_hLib != NULL; }

    HRESULT CreateInArchive(const GUID& formatId, IInArchive** archive) {
        if (!m_createObject) return E_FAIL;
        return m_createObject(&formatId, &IID_IInArchive, (void**)archive);
    }

    HRESULT CreateOutArchive(const GUID& formatId, IOutArchive** archive) {
        if (!m_createObject) return E_FAIL;
        return m_createObject(&formatId, &IID_IOutArchive, (void**)archive);
    }

    const GUID* GetFormatIdForExtension(const std::wstring& ext) {
        for (const auto& fmt : g_archiveFormats) {
            if (_wcsicmp(ext.c_str(), fmt.ext) == 0) {
                return fmt.formatId;
            }
        }
        return nullptr;
    }

private:
    C7ZipLib() : m_hLib(NULL), m_createObject(nullptr) {
        Init7ZipPaths();
        if (!g_7zDllPath.empty()) {
            m_hLib = LoadLibraryW(g_7zDllPath.c_str());
            if (m_hLib) {
                m_createObject = (Func_CreateObject)GetProcAddress(m_hLib, "CreateObject");
            }
        }
    }
    ~C7ZipLib() {
        if (m_hLib) FreeLibrary(m_hLib);
    }

    HMODULE m_hLib;
    Func_CreateObject m_createObject;
};

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
    // Sub commands will be created dynamically based on file type
}

bool CExplorerCommand::IsArchiveFile(const std::wstring& path)
{
    std::wstring ext = PathFindExtensionW(path.c_str());
    for (const auto& fmt : g_archiveFormats) {
        if (_wcsicmp(ext.c_str(), fmt.ext) == 0) {
            return true;
        }
    }
    return false;
}

void CExplorerCommand::GetSelectedItems(IShellItemArray* psiItemArray)
{
    if (!psiItemArray) return;

    DWORD count = 0;
    psiItemArray->GetCount(&count);
    if (count > 0) {
        IShellItem* psi = nullptr;
        if (SUCCEEDED(psiItemArray->GetItemAt(0, &psi))) {
            PWSTR pszPath = nullptr;
            if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath))) {
                m_selectedPath = pszPath;
                m_isArchive = IsArchiveFile(m_selectedPath);
                CoTaskMemFree(pszPath);
            }
            psi->Release();
        }
    }
}

// IUnknown
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

// IExplorerCommand
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
    Init7ZipPaths();
    wchar_t iconPath[MAX_PATH];
    StringCchPrintfW(iconPath, MAX_PATH, L"%s,0", g_7zFMPath.c_str());
    return SHStrDupW(iconPath, ppszIcon);
}

IFACEMETHODIMP CExplorerCommand::GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip)
{
    *ppszInfotip = NULL;
    return E_NOTIMPL;
}

IFACEMETHODIMP CExplorerCommand::GetCanonicalName(GUID* pguidCommandName)
{
    // Each command type needs a unique GUID or Windows will deduplicate them
    // Only root command returns a valid GUID, subcommands return E_NOTIMPL
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
    *pCmdState = ECS_ENABLED;
    return S_OK;
}

// Helper: Get file name without path
static std::wstring GetFileName(const std::wstring& path)
{
    size_t pos = path.rfind(L'\\');
    return (pos != std::wstring::npos) ? path.substr(pos + 1) : path;
}

// Helper: Get file name without extension
static std::wstring GetFileNameWithoutExt(const std::wstring& path)
{
    std::wstring name = GetFileName(path);
    size_t pos = name.rfind(L'.');
    return (pos != std::wstring::npos) ? name.substr(0, pos) : name;
}

// Helper: Get parent directory
static std::wstring GetParentDir(const std::wstring& path)
{
    size_t pos = path.rfind(L'\\');
    return (pos != std::wstring::npos) ? path.substr(0, pos) : path;
}

// Helper: Show password input dialog
static std::wstring ShowPasswordDialog()
{
    // Simple input box using Windows API
    wchar_t password[256] = {};

    // Create a simple dialog - for now just return empty
    // In production, use a proper dialog or TaskDialog
    const auto& strings = GetLocalizedStrings();

    // Use a simple approach: prompt via a basic dialog
    // For better UX, could implement a proper dialog resource
    return L"";  // User needs to add password manually to passwords.txt for now
}

// Show password result dialog after successful extraction
static void ShowPasswordResultDialog(const std::wstring& password)
{
    const auto& strings = GetLocalizedStrings();
    wchar_t message[1024];
    StringCchPrintfW(message, 1024, strings.passwordMessage, password.c_str());
    MessageBoxW(NULL, message, strings.passwordTitle, MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND);
}

// Extract archive using 7-Zip SDK with password support
bool CExplorerCommand::ExtractArchive(const std::wstring& archivePath, const std::wstring& outDir)
{
    C7ZipLib& lib = C7ZipLib::Instance();
    if (!lib.IsLoaded()) return false;

    // Get format ID from extension
    std::wstring ext = PathFindExtensionW(archivePath.c_str());
    const GUID* formatId = lib.GetFormatIdForExtension(ext);

    // For formats without specific GUID, try 7z format (it auto-detects)
    if (!formatId) formatId = &CLSID_CFormat7z;

    // Get stored passwords
    CPasswordManager& pwdMgr = CPasswordManager::Instance();
    std::vector<std::wstring> passwords = pwdMgr.GetPasswords();

    // Add empty password at the beginning (for non-encrypted archives)
    passwords.insert(passwords.begin(), L"");

    std::wstring successPassword;
    bool success = false;

    for (const auto& password : passwords) {
        IInArchive* archive = nullptr;
        if (FAILED(lib.CreateInArchive(*formatId, &archive))) continue;

        CFileInStream* inStream = new CFileInStream();
        if (!inStream->Open(archivePath.c_str())) {
            inStream->Release();
            archive->Release();
            continue;
        }

        UINT64 maxCheckStartPosition = 1 << 22;  // 4MB
        if (FAILED(archive->Open(inStream, &maxCheckStartPosition, nullptr))) {
            inStream->Release();
            archive->Release();
            continue;
        }

        CExtractCallback* extractCallback = new CExtractCallback(outDir, archive, password);
        HRESULT hr = archive->Extract(nullptr, (UINT32)-1, 0,
            static_cast<IArchiveExtractCallback*>(extractCallback));

        bool wasPasswordUsed = extractCallback->WasPasswordUsed();
        extractCallback->Release();
        archive->Close();
        archive->Release();
        inStream->Release();

        if (SUCCEEDED(hr)) {
            success = true;
            if (wasPasswordUsed && !password.empty()) {
                successPassword = password;
            }
            break;
        }
    }

    // Update password count and show dialog if a stored password was used
    if (success && !successPassword.empty()) {
        pwdMgr.IncrementCount(successPassword);
        ShowPasswordResultDialog(successPassword);
    }

    return success;
}

// Compress files using 7-Zip SDK
bool CExplorerCommand::CompressFiles(const std::wstring& srcPath, const std::wstring& archivePath, const GUID& formatId)
{
    C7ZipLib& lib = C7ZipLib::Instance();
    if (!lib.IsLoaded()) return false;

    IOutArchive* archive = nullptr;
    if (FAILED(lib.CreateOutArchive(formatId, &archive))) return false;

    CFileOutStream* outStream = new CFileOutStream();
    if (!outStream->Create(archivePath.c_str())) {
        outStream->Release();
        archive->Release();
        return false;
    }

    std::wstring itemName = GetFileName(srcPath);
    CUpdateCallback* updateCallback = new CUpdateCallback(srcPath, itemName);

    HRESULT hr = archive->UpdateItems(outStream, updateCallback->GetItemCount(), updateCallback);

    updateCallback->Release();
    outStream->Release();
    archive->Release();

    return SUCCEEDED(hr);
}

// Check if files exist in destination directory that would be overwritten
static bool CheckOverwriteNeededForArchive(const std::wstring& archivePath, const std::wstring& outDir)
{
    C7ZipLib& lib = C7ZipLib::Instance();
    if (!lib.IsLoaded()) return false;

    std::wstring ext = PathFindExtensionW(archivePath.c_str());
    const GUID* formatId = lib.GetFormatIdForExtension(ext);
    if (!formatId) formatId = &CLSID_CFormat7z;

    IInArchive* archive = nullptr;
    if (FAILED(lib.CreateInArchive(*formatId, &archive))) return false;

    CFileInStream* inStream = new CFileInStream();
    if (!inStream->Open(archivePath.c_str())) {
        inStream->Release();
        archive->Release();
        return false;
    }

    UINT64 maxCheckStartPosition = 1 << 22;
    if (FAILED(archive->Open(inStream, &maxCheckStartPosition, nullptr))) {
        inStream->Release();
        archive->Release();
        return false;
    }

    bool needOverwrite = false;
    UINT32 numItems = 0;
    if (SUCCEEDED(archive->GetNumberOfItems(&numItems))) {
        for (UINT32 i = 0; i < numItems && !needOverwrite; i++) {
            PROPVARIANT prop;
            PropVariantInit(&prop);
            archive->GetProperty(i, kpidPath, &prop);
            if (prop.vt == VT_BSTR && prop.bstrVal) {
                std::wstring fullPath = outDir + L"\\" + prop.bstrVal;
                if (GetFileAttributesW(fullPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    needOverwrite = true;
                }
            }
            PropVariantClear(&prop);
        }
    }

    archive->Close();
    archive->Release();
    inStream->Release();

    return needOverwrite;
}

// Show overwrite confirmation dialog
static bool ConfirmOverwrite()
{
    const auto& strings = GetLocalizedStrings();
    int result = MessageBoxW(NULL, strings.overwriteMessage, strings.overwriteTitle,
                             MB_YESNO | MB_ICONQUESTION | MB_SETFOREGROUND);
    return (result == IDYES);
}

IFACEMETHODIMP CExplorerCommand::Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc)
{
    GetSelectedItems(psiItemArray);
    if (m_selectedPath.empty()) return E_FAIL;

    bool success = false;

    switch (m_type) {
        case CommandType::ExtractHere:
            // Extract to parent directory using SDK
            {
                std::wstring outDir = GetParentDir(m_selectedPath);
                if (CheckOverwriteNeededForArchive(m_selectedPath, outDir)) {
                    if (!ConfirmOverwrite()) {
                        return S_OK;  // User cancelled
                    }
                }
                success = ExtractArchive(m_selectedPath, outDir);
            }
            break;

        case CommandType::ExtractTo:
            // Extract to subfolder using SDK
            {
                std::wstring subDir = GetParentDir(m_selectedPath) + L"\\" + GetFileNameWithoutExt(m_selectedPath);
                if (GetFileAttributesW(subDir.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    // Subfolder exists, check for overwrites
                    if (CheckOverwriteNeededForArchive(m_selectedPath, subDir)) {
                        if (!ConfirmOverwrite()) {
                            return S_OK;  // User cancelled
                        }
                    }
                }
                CreateDirectoryW(subDir.c_str(), NULL);
                success = ExtractArchive(m_selectedPath, subDir);
            }
            break;

        case CommandType::AddTo7z:
            // Compress to .7z using SDK
            success = CompressFiles(m_selectedPath, m_selectedPath + L".7z", CLSID_CFormat7z);
            break;

        case CommandType::AddToZip:
            // Compress to .zip using SDK
            success = CompressFiles(m_selectedPath, m_selectedPath + L".zip", CLSID_CFormatZip);
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

    // Clear and rebuild sub commands
    for (auto cmd : m_subCommands) {
        cmd->Release();
    }
    m_subCommands.clear();

    // Add all possible sub-commands - GetState will hide inappropriate ones
    m_subCommands.push_back(new CExplorerCommand(CommandType::ExtractHere));
    m_subCommands.push_back(new CExplorerCommand(CommandType::ExtractTo));
    m_subCommands.push_back(new CExplorerCommand(CommandType::AddTo7z));
    m_subCommands.push_back(new CExplorerCommand(CommandType::AddToZip));

    m_enumIndex = 0;
    this->AddRef();
    *ppEnum = this;
    return S_OK;
}

// IEnumExplorerCommand
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
