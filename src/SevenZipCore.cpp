// SevenZipCore.cpp - 7-Zip functionality wrapper implementation
#include "SevenZipCore.h"

#include <Windows.h>
#include <PropIdl.h>
#include <shlwapi.h>

// 7-Zip headers (GUIDs are instantiated in GuidInit.cpp)
#include "Common/Common.h"
#include "7zip/Archive/IArchive.h"
#include "7zip/IPassword.h"
#include "7zip/Common/FileStreams.h"
#include "7zip/PropID.h"
#include "Common/MyCom.h"
#include "Common/MyString.h"

// Archive format GUIDs
// {23170F69-40C1-278A-1000-000110070000} - 7z
static const GUID CLSID_CFormat7z =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x07, 0x00, 0x00 } };

// {23170F69-40C1-278A-1000-000110010000} - Zip
static const GUID CLSID_CFormatZip =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x01, 0x00, 0x00 } };

// {23170F69-40C1-278A-1000-000110030000} - Rar
static const GUID CLSID_CFormatRar =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x03, 0x00, 0x00 } };

// {23170F69-40C1-278A-1000-000110EE0000} - Tar
static const GUID CLSID_CFormatTar =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0xEE, 0x00, 0x00 } };

// {23170F69-40C1-278A-1000-000110EF0000} - GZip
static const GUID CLSID_CFormatGZip =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0xEF, 0x00, 0x00 } };

// {23170F69-40C1-278A-1000-000110020000} - BZip2
static const GUID CLSID_CFormatBZip2 =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x02, 0x00, 0x00 } };

// {23170F69-40C1-278A-1000-000110E70000} - Iso
static const GUID CLSID_CFormatIso =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0xE7, 0x00, 0x00 } };

// {23170F69-40C1-278A-1000-000110080000} - Cab
static const GUID CLSID_CFormatCab =
    { 0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x08, 0x00, 0x00 } };

// External function from 7-Zip to create archive objects
STDAPI CreateObject(const GUID *clsid, const GUID *iid, void **outObject);

//////////////////////////////////////////////////////////////////////////////
// Helper: Create directory recursively
//////////////////////////////////////////////////////////////////////////////

static void CreateDirectoryRecursive(const std::wstring& path) {
    size_t pos = 0;
    while ((pos = path.find(L'\\', pos + 1)) != std::wstring::npos) {
        CreateDirectoryW(path.substr(0, pos).c_str(), NULL);
    }
    CreateDirectoryW(path.c_str(), NULL);
}

//////////////////////////////////////////////////////////////////////////////
// Simple Output File Stream (minimal implementation for extraction)
//////////////////////////////////////////////////////////////////////////////

class CSimpleOutFileStream :
    public ISequentialOutStream,
    public CMyUnknownImp
{
public:
    CSimpleOutFileStream() : m_hFile(INVALID_HANDLE_VALUE), m_refCount(0) {}
    virtual ~CSimpleOutFileStream() { Close(); }

    bool Create(const wchar_t* path, bool createAlways = true) {
        // Create parent directory if needed
        std::wstring dir = path;
        size_t pos = dir.rfind(L'\\');
        if (pos != std::wstring::npos) {
            dir = dir.substr(0, pos);
            CreateDirectoryRecursive(dir);
        }

        m_hFile = CreateFileW(path, GENERIC_WRITE, 0, NULL,
                              createAlways ? CREATE_ALWAYS : CREATE_NEW,
                              FILE_ATTRIBUTE_NORMAL, NULL);
        return m_hFile != INVALID_HANDLE_VALUE;
    }

    void Close() {
        if (m_hFile != INVALID_HANDLE_VALUE) {
            CloseHandle(m_hFile);
            m_hFile = INVALID_HANDLE_VALUE;
        }
    }

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID iid, void **outObject) {
        if (iid == IID_IUnknown || iid == IID_ISequentialOutStream) {
            *outObject = this;
            AddRef();
            return S_OK;
        }
        *outObject = NULL;
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() { return ++m_refCount; }
    STDMETHOD_(ULONG, Release)() {
        ULONG res = --m_refCount;
        if (res == 0) delete this;
        return res;
    }

    STDMETHOD(Write)(const void *data, UInt32 size, UInt32 *processedSize) {
        DWORD written = 0;
        if (!WriteFile(m_hFile, data, size, &written, NULL)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (processedSize) *processedSize = written;
        return S_OK;
    }

private:
    HANDLE m_hFile;
    ULONG m_refCount;
};

//////////////////////////////////////////////////////////////////////////////
// Extract Callback
//////////////////////////////////////////////////////////////////////////////

class CExtractCallback :
    public IArchiveExtractCallback,
    public ICryptoGetTextPassword,
    public CMyUnknownImp
{
public:
    CExtractCallback(IInArchive* archive, const std::wstring& outDir,
                     const std::wstring& password, ProgressCallback progress)
        : m_archive(archive)
        , m_outDir(outDir)
        , m_password(password)
        , m_progress(progress)
        , m_total(0)
        , m_completed(0)
        , m_passwordWasRequested(false)
        , m_refCount(0)
    {}

    bool WasPasswordRequested() const { return m_passwordWasRequested; }

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID iid, void **outObject) {
        if (iid == IID_IUnknown) {
            *outObject = static_cast<IArchiveExtractCallback*>(this);
        } else if (iid == IID_IArchiveExtractCallback) {
            *outObject = static_cast<IArchiveExtractCallback*>(this);
        } else if (iid == IID_ICryptoGetTextPassword) {
            *outObject = static_cast<ICryptoGetTextPassword*>(this);
        } else {
            *outObject = NULL;
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }

    STDMETHOD_(ULONG, AddRef)() { return ++m_refCount; }
    STDMETHOD_(ULONG, Release)() {
        ULONG res = --m_refCount;
        if (res == 0) delete this;
        return res;
    }

    // IProgress
    STDMETHOD(SetTotal)(UInt64 total) {
        m_total = total;
        return S_OK;
    }

    STDMETHOD(SetCompleted)(const UInt64 *completeValue) {
        if (completeValue) {
            m_completed = *completeValue;
            if (m_progress && !m_progress(m_completed, m_total)) {
                return E_ABORT;
            }
        }
        return S_OK;
    }

    // IArchiveExtractCallback
    STDMETHOD(GetStream)(UInt32 index, ISequentialOutStream **outStream, Int32 askExtractMode) {
        *outStream = NULL;
        if (askExtractMode != NArchive::NExtract::NAskMode::kExtract) {
            return S_OK;
        }

        // Get item path
        PROPVARIANT prop;
        PropVariantInit(&prop);
        m_archive->GetProperty(index, kpidPath, &prop);
        std::wstring itemPath = (prop.vt == VT_BSTR) ? prop.bstrVal : L"";
        PropVariantClear(&prop);

        // Check if directory
        PropVariantInit(&prop);
        m_archive->GetProperty(index, kpidIsDir, &prop);
        bool isDir = (prop.vt == VT_BOOL && prop.boolVal != VARIANT_FALSE);
        PropVariantClear(&prop);

        std::wstring fullPath = m_outDir;
        if (!fullPath.empty() && fullPath.back() != L'\\') {
            fullPath += L'\\';
        }
        fullPath += itemPath;

        if (isDir) {
            CreateDirectoryRecursive(fullPath);
            return S_OK;
        }

        // Create output stream
        CSimpleOutFileStream* stream = new CSimpleOutFileStream();
        stream->AddRef();
        if (!stream->Create(fullPath.c_str())) {
            stream->Release();
            return HRESULT_FROM_WIN32(GetLastError());
        }
        *outStream = stream;
        return S_OK;
    }

    STDMETHOD(PrepareOperation)(Int32 askExtractMode) {
        return S_OK;
    }

    STDMETHOD(SetOperationResult)(Int32 opRes) {
        return S_OK;
    }

    // ICryptoGetTextPassword
    STDMETHOD(CryptoGetTextPassword)(BSTR *password) {
        m_passwordWasRequested = true;
        *password = SysAllocString(m_password.c_str());
        return S_OK;
    }

private:
    IInArchive* m_archive;
    std::wstring m_outDir;
    std::wstring m_password;
    ProgressCallback m_progress;
    UInt64 m_total;
    UInt64 m_completed;
    bool m_passwordWasRequested;
    ULONG m_refCount;
};

//////////////////////////////////////////////////////////////////////////////
// Simple Input File Stream (minimal implementation for compression)
//////////////////////////////////////////////////////////////////////////////

class CSimpleInFileStream :
    public ISequentialInStream,
    public CMyUnknownImp
{
public:
    CSimpleInFileStream() : m_hFile(INVALID_HANDLE_VALUE), m_refCount(0) {}
    virtual ~CSimpleInFileStream() { Close(); }

    bool Open(const wchar_t* path) {
        m_hFile = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ,
                              NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        return m_hFile != INVALID_HANDLE_VALUE;
    }

    void Close() {
        if (m_hFile != INVALID_HANDLE_VALUE) {
            CloseHandle(m_hFile);
            m_hFile = INVALID_HANDLE_VALUE;
        }
    }

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID iid, void **outObject) {
        if (iid == IID_IUnknown || iid == IID_ISequentialInStream) {
            *outObject = this;
            AddRef();
            return S_OK;
        }
        *outObject = NULL;
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() { return ++m_refCount; }
    STDMETHOD_(ULONG, Release)() {
        ULONG res = --m_refCount;
        if (res == 0) delete this;
        return res;
    }

    STDMETHOD(Read)(void *data, UInt32 size, UInt32 *processedSize) {
        DWORD read = 0;
        if (!ReadFile(m_hFile, data, size, &read, NULL)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (processedSize) *processedSize = read;
        return S_OK;
    }

private:
    HANDLE m_hFile;
    ULONG m_refCount;
};

//////////////////////////////////////////////////////////////////////////////
// Full Input File Stream (for archive opening)
//////////////////////////////////////////////////////////////////////////////

class CFullInFileStream :
    public IInStream,
    public IStreamGetSize,
    public CMyUnknownImp
{
public:
    CFullInFileStream() : m_hFile(INVALID_HANDLE_VALUE), m_refCount(0) {}
    virtual ~CFullInFileStream() { Close(); }

    bool Open(const wchar_t* path) {
        m_hFile = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ,
                              NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        return m_hFile != INVALID_HANDLE_VALUE;
    }

    void Close() {
        if (m_hFile != INVALID_HANDLE_VALUE) {
            CloseHandle(m_hFile);
            m_hFile = INVALID_HANDLE_VALUE;
        }
    }

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID iid, void **outObject) {
        if (iid == IID_IUnknown) {
            *outObject = static_cast<IInStream*>(this);
            AddRef();
            return S_OK;
        }
        if (iid == IID_IInStream || iid == IID_ISequentialInStream) {
            *outObject = static_cast<IInStream*>(this);
            AddRef();
            return S_OK;
        }
        if (iid == IID_IStreamGetSize) {
            *outObject = static_cast<IStreamGetSize*>(this);
            AddRef();
            return S_OK;
        }
        *outObject = NULL;
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() { return ++m_refCount; }
    STDMETHOD_(ULONG, Release)() {
        ULONG res = --m_refCount;
        if (res == 0) delete this;
        return res;
    }

    STDMETHOD(Read)(void *data, UInt32 size, UInt32 *processedSize) {
        DWORD read = 0;
        if (!ReadFile(m_hFile, data, size, &read, NULL)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (processedSize) *processedSize = read;
        return S_OK;
    }

    STDMETHOD(Seek)(Int64 offset, UInt32 seekOrigin, UInt64 *newPosition) {
        LARGE_INTEGER li;
        li.QuadPart = offset;
        LARGE_INTEGER newPos;
        if (!SetFilePointerEx(m_hFile, li, &newPos, seekOrigin)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (newPosition) *newPosition = newPos.QuadPart;
        return S_OK;
    }

    STDMETHOD(GetSize)(UInt64 *size) {
        LARGE_INTEGER li;
        if (!GetFileSizeEx(m_hFile, &li)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        *size = li.QuadPart;
        return S_OK;
    }

private:
    HANDLE m_hFile;
    ULONG m_refCount;
};

//////////////////////////////////////////////////////////////////////////////
// Full Output File Stream (for archive creation)
//////////////////////////////////////////////////////////////////////////////

class CFullOutFileStream :
    public IOutStream,
    public CMyUnknownImp
{
public:
    CFullOutFileStream() : m_hFile(INVALID_HANDLE_VALUE), m_refCount(0) {}
    virtual ~CFullOutFileStream() { Close(); }

    bool Create(const wchar_t* path, bool createAlways = true) {
        // Create parent directory if needed
        std::wstring dir = path;
        size_t pos = dir.rfind(L'\\');
        if (pos != std::wstring::npos) {
            dir = dir.substr(0, pos);
            CreateDirectoryRecursive(dir);
        }

        m_hFile = CreateFileW(path, GENERIC_WRITE, 0, NULL,
                              createAlways ? CREATE_ALWAYS : CREATE_NEW,
                              FILE_ATTRIBUTE_NORMAL, NULL);
        return m_hFile != INVALID_HANDLE_VALUE;
    }

    void Close() {
        if (m_hFile != INVALID_HANDLE_VALUE) {
            CloseHandle(m_hFile);
            m_hFile = INVALID_HANDLE_VALUE;
        }
    }

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID iid, void **outObject) {
        if (iid == IID_IUnknown) {
            *outObject = static_cast<IOutStream*>(this);
            AddRef();
            return S_OK;
        }
        if (iid == IID_IOutStream || iid == IID_ISequentialOutStream) {
            *outObject = static_cast<IOutStream*>(this);
            AddRef();
            return S_OK;
        }
        *outObject = NULL;
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() { return ++m_refCount; }
    STDMETHOD_(ULONG, Release)() {
        ULONG res = --m_refCount;
        if (res == 0) delete this;
        return res;
    }

    STDMETHOD(Write)(const void *data, UInt32 size, UInt32 *processedSize) {
        DWORD written = 0;
        if (!WriteFile(m_hFile, data, size, &written, NULL)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (processedSize) *processedSize = written;
        return S_OK;
    }

    STDMETHOD(Seek)(Int64 offset, UInt32 seekOrigin, UInt64 *newPosition) {
        LARGE_INTEGER li;
        li.QuadPart = offset;
        LARGE_INTEGER newPos;
        if (!SetFilePointerEx(m_hFile, li, &newPos, seekOrigin)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (newPosition) *newPosition = newPos.QuadPart;
        return S_OK;
    }

    STDMETHOD(SetSize)(UInt64 newSize) {
        LARGE_INTEGER li;
        li.QuadPart = newSize;
        if (!SetFilePointerEx(m_hFile, li, NULL, FILE_BEGIN)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (!SetEndOfFile(m_hFile)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        return S_OK;
    }

private:
    HANDLE m_hFile;
    ULONG m_refCount;
};

//////////////////////////////////////////////////////////////////////////////
// Update Callback for Compression
//////////////////////////////////////////////////////////////////////////////

class CUpdateCallback :
    public IArchiveUpdateCallback,
    public ICryptoGetTextPassword2,
    public CMyUnknownImp
{
public:
    struct FileItem {
        std::wstring fullPath;
        std::wstring relativePath;
        bool isDir;
        UInt64 size;
        FILETIME mtime;
        DWORD attrib;
    };

    CUpdateCallback(const std::vector<std::wstring>& srcPaths, ProgressCallback progress)
        : m_progress(progress)
        , m_total(0)
        , m_completed(0)
        , m_refCount(0)
    {
        // Enumerate all files
        for (const auto& srcPath : srcPaths) {
            std::wstring name = GetFileName(srcPath);
            DWORD attrs = GetFileAttributesW(srcPath.c_str());
            bool isDir = (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY));

            if (isDir) {
                // Add directory and its contents
                FileItem item;
                item.fullPath = srcPath;
                item.relativePath = name;
                item.isDir = true;
                item.size = 0;
                item.attrib = attrs;
                GetFileTimeHelper(srcPath, item.mtime);
                m_files.push_back(item);
                EnumerateDirectory(srcPath, name);
            } else {
                // Single file
                FileItem item;
                item.fullPath = srcPath;
                item.relativePath = name;
                item.isDir = false;
                item.attrib = attrs;
                GetFileSizeHelper(srcPath, item.size);
                GetFileTimeHelper(srcPath, item.mtime);
                m_files.push_back(item);
                m_total += item.size;
            }
        }
    }

    UInt32 GetItemCount() const { return (UInt32)m_files.size(); }

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID iid, void **outObject) {
        if (iid == IID_IUnknown) {
            *outObject = static_cast<IArchiveUpdateCallback*>(this);
        } else if (iid == IID_IArchiveUpdateCallback) {
            *outObject = static_cast<IArchiveUpdateCallback*>(this);
        } else if (iid == IID_ICryptoGetTextPassword2) {
            *outObject = static_cast<ICryptoGetTextPassword2*>(this);
        } else {
            *outObject = NULL;
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }

    STDMETHOD_(ULONG, AddRef)() { return ++m_refCount; }
    STDMETHOD_(ULONG, Release)() {
        ULONG res = --m_refCount;
        if (res == 0) delete this;
        return res;
    }

    // IProgress
    STDMETHOD(SetTotal)(UInt64 total) {
        m_total = total;
        return S_OK;
    }

    STDMETHOD(SetCompleted)(const UInt64 *completeValue) {
        if (completeValue) {
            m_completed = *completeValue;
            if (m_progress && !m_progress(m_completed, m_total)) {
                return E_ABORT;
            }
        }
        return S_OK;
    }

    // IArchiveUpdateCallback
    STDMETHOD(GetUpdateItemInfo)(UInt32 index, Int32 *newData, Int32 *newProps, UInt32 *indexInArchive) {
        if (newData) *newData = 1;      // New data
        if (newProps) *newProps = 1;    // New properties
        if (indexInArchive) *indexInArchive = (UInt32)-1;  // Not in archive
        return S_OK;
    }

    STDMETHOD(GetProperty)(UInt32 index, PROPID propID, PROPVARIANT *value) {
        PropVariantInit(value);
        if (index >= m_files.size()) return E_INVALIDARG;

        const FileItem& item = m_files[index];

        switch (propID) {
            case kpidPath:
                value->vt = VT_BSTR;
                value->bstrVal = SysAllocString(item.relativePath.c_str());
                break;
            case kpidIsDir:
                value->vt = VT_BOOL;
                value->boolVal = item.isDir ? VARIANT_TRUE : VARIANT_FALSE;
                break;
            case kpidSize:
                if (!item.isDir) {
                    value->vt = VT_UI8;
                    value->uhVal.QuadPart = item.size;
                }
                break;
            case kpidAttrib:
                value->vt = VT_UI4;
                value->ulVal = item.attrib;
                break;
            case kpidMTime:
                value->vt = VT_FILETIME;
                value->filetime = item.mtime;
                break;
        }
        return S_OK;
    }

    STDMETHOD(GetStream)(UInt32 index, ISequentialInStream **inStream) {
        *inStream = NULL;
        if (index >= m_files.size()) return E_INVALIDARG;

        const FileItem& item = m_files[index];
        if (item.isDir) return S_OK;

        CSimpleInFileStream* stream = new CSimpleInFileStream();
        stream->AddRef();
        if (!stream->Open(item.fullPath.c_str())) {
            stream->Release();
            return HRESULT_FROM_WIN32(GetLastError());
        }
        *inStream = stream;
        return S_OK;
    }

    STDMETHOD(SetOperationResult)(Int32 operationResult) {
        return S_OK;
    }

    // ICryptoGetTextPassword2
    STDMETHOD(CryptoGetTextPassword2)(Int32 *passwordIsDefined, BSTR *password) {
        *passwordIsDefined = 0;
        *password = SysAllocString(L"");
        return S_OK;
    }

private:
    std::vector<FileItem> m_files;
    ProgressCallback m_progress;
    UInt64 m_total;
    UInt64 m_completed;
    ULONG m_refCount;

    std::wstring GetFileName(const std::wstring& path) {
        size_t pos = path.rfind(L'\\');
        return (pos != std::wstring::npos) ? path.substr(pos + 1) : path;
    }

    void GetFileSizeHelper(const std::wstring& path, UInt64& size) {
        WIN32_FILE_ATTRIBUTE_DATA fad;
        if (GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &fad)) {
            size = ((UInt64)fad.nFileSizeHigh << 32) | fad.nFileSizeLow;
        } else {
            size = 0;
        }
    }

    void GetFileTimeHelper(const std::wstring& path, FILETIME& mtime) {
        WIN32_FILE_ATTRIBUTE_DATA fad;
        if (GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &fad)) {
            mtime = fad.ftLastWriteTime;
        } else {
            GetSystemTimeAsFileTime(&mtime);
        }
    }

    void EnumerateDirectory(const std::wstring& basePath, const std::wstring& relBasePath) {
        std::wstring searchPath = basePath + L"\\*";
        WIN32_FIND_DATAW fd;
        HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
        if (hFind == INVALID_HANDLE_VALUE) return;

        do {
            if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) {
                continue;
            }

            std::wstring fullPath = basePath + L"\\" + fd.cFileName;
            std::wstring relPath = relBasePath + L"\\" + fd.cFileName;

            FileItem item;
            item.fullPath = fullPath;
            item.relativePath = relPath;
            item.isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            item.attrib = fd.dwFileAttributes;
            item.mtime = fd.ftLastWriteTime;

            if (item.isDir) {
                item.size = 0;
                m_files.push_back(item);
                EnumerateDirectory(fullPath, relPath);
            } else {
                item.size = ((UInt64)fd.nFileSizeHigh << 32) | fd.nFileSizeLow;
                m_files.push_back(item);
                m_total += item.size;
            }
        } while (FindNextFileW(hFind, &fd));

        FindClose(hFind);
    }
};

//////////////////////////////////////////////////////////////////////////////
// SevenZipCore Implementation
//////////////////////////////////////////////////////////////////////////////

SevenZipCore& SevenZipCore::Instance() {
    static SevenZipCore instance;
    return instance;
}

SevenZipCore::SevenZipCore() {
    InitFormats();
}

SevenZipCore::~SevenZipCore() {
    CloseArchive();
}

void SevenZipCore::InitFormats() {
    m_formats = {
        { L"7z",    L".7z",   CLSID_CFormat7z,    true  },
        { L"Zip",   L".zip",  CLSID_CFormatZip,   true  },
        { L"Rar",   L".rar",  CLSID_CFormatRar,   false },
        { L"Tar",   L".tar",  CLSID_CFormatTar,   true  },
        { L"GZip",  L".gz",   CLSID_CFormatGZip,  true  },
        { L"BZip2", L".bz2",  CLSID_CFormatBZip2, true  },
        { L"Iso",   L".iso",  CLSID_CFormatIso,   false },
        { L"Cab",   L".cab",  CLSID_CFormatCab,   false },
    };
}

IInArchive* SevenZipCore::CreateInArchive(const GUID& formatId) {
    IInArchive* archive = nullptr;
    HRESULT hr = CreateObject(&formatId, &IID_IInArchive, (void**)&archive);
    if (FAILED(hr)) {
        return nullptr;
    }
    return archive;
}

IOutArchive* SevenZipCore::CreateOutArchive(const GUID& formatId) {
    IOutArchive* archive = nullptr;
    HRESULT hr = CreateObject(&formatId, &IID_IOutArchive, (void**)&archive);
    if (FAILED(hr)) {
        return nullptr;
    }
    return archive;
}

const GUID* SevenZipCore::GetFormatForExtension(const std::wstring& ext) {
    for (const auto& fmt : m_formats) {
        if (_wcsicmp(ext.c_str(), fmt.extension.c_str()) == 0) {
            return &fmt.classId;
        }
    }
    // Handle compound extensions
    if (_wcsicmp(ext.c_str(), L".tgz") == 0 ||
        _wcsicmp(ext.c_str(), L".tar.gz") == 0) {
        return &CLSID_CFormatGZip;
    }
    if (_wcsicmp(ext.c_str(), L".tbz2") == 0 ||
        _wcsicmp(ext.c_str(), L".tar.bz2") == 0) {
        return &CLSID_CFormatBZip2;
    }
    return nullptr;
}

const GUID* SevenZipCore::DetectFormat(const std::wstring& path) {
    // First try by extension
    const wchar_t* ext = PathFindExtensionW(path.c_str());
    if (ext && *ext) {
        const GUID* guid = GetFormatForExtension(ext);
        if (guid) return guid;
    }

    // TODO: Add signature-based detection
    return &CLSID_CFormat7z;  // Default to 7z
}

bool SevenZipCore::OpenArchive(const std::wstring& path) {
    CloseArchive();

    // Detect format
    const GUID* formatId = DetectFormat(path);
    if (!formatId) {
        return false;
    }

    // Create archive handler
    m_archive = CreateInArchive(*formatId);
    if (!m_archive) {
        return false;
    }

    // Open file stream
    CFullInFileStream* inStream = new CFullInFileStream();
    inStream->AddRef();
    if (!inStream->Open(path.c_str())) {
        inStream->Release();
        m_archive->Release();
        m_archive = nullptr;
        return false;
    }
    m_inStream = inStream;

    // Open archive
    UInt64 maxCheckStartPosition = 1 << 22;
    HRESULT hr = m_archive->Open(m_inStream, &maxCheckStartPosition, nullptr);
    if (FAILED(hr)) {
        m_inStream->Release();
        m_inStream = nullptr;
        m_archive->Release();
        m_archive = nullptr;
        return false;
    }

    m_currentPath = path;
    m_needsPassword = false;

    // Check if any item is encrypted
    UInt32 numItems = 0;
    m_archive->GetNumberOfItems(&numItems);
    for (UInt32 i = 0; i < numItems; i++) {
        PROPVARIANT prop;
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidEncrypted, &prop);
        if (prop.vt == VT_BOOL && prop.boolVal != VARIANT_FALSE) {
            m_needsPassword = true;
            PropVariantClear(&prop);
            break;
        }
        PropVariantClear(&prop);
    }

    return true;
}

void SevenZipCore::CloseArchive() {
    if (m_archive) {
        m_archive->Close();
        m_archive->Release();
        m_archive = nullptr;
    }
    if (m_inStream) {
        m_inStream->Release();
        m_inStream = nullptr;
    }
    m_currentPath.clear();
    m_needsPassword = false;
}

uint32_t SevenZipCore::GetItemCount() {
    if (!m_archive) return 0;
    UInt32 count = 0;
    m_archive->GetNumberOfItems(&count);
    return count;
}

std::vector<ArchiveItem> SevenZipCore::GetItems() {
    std::vector<ArchiveItem> items;
    if (!m_archive) return items;

    UInt32 numItems = 0;
    m_archive->GetNumberOfItems(&numItems);

    for (UInt32 i = 0; i < numItems; i++) {
        ArchiveItem item;
        PROPVARIANT prop;

        // Path
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidPath, &prop);
        item.path = (prop.vt == VT_BSTR) ? prop.bstrVal : L"";
        PropVariantClear(&prop);

        // Size
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidSize, &prop);
        item.size = (prop.vt == VT_UI8) ? prop.uhVal.QuadPart : 0;
        PropVariantClear(&prop);

        // Packed size
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidPackSize, &prop);
        item.packedSize = (prop.vt == VT_UI8) ? prop.uhVal.QuadPart : 0;
        PropVariantClear(&prop);

        // Is directory
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidIsDir, &prop);
        item.isDir = (prop.vt == VT_BOOL && prop.boolVal != VARIANT_FALSE);
        PropVariantClear(&prop);

        // Is encrypted
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidEncrypted, &prop);
        item.isEncrypted = (prop.vt == VT_BOOL && prop.boolVal != VARIANT_FALSE);
        PropVariantClear(&prop);

        // Modification time
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidMTime, &prop);
        if (prop.vt == VT_FILETIME) {
            item.mtime = prop.filetime;
        } else {
            memset(&item.mtime, 0, sizeof(item.mtime));
        }
        PropVariantClear(&prop);

        items.push_back(item);
    }

    return items;
}

bool SevenZipCore::TestPassword(const std::wstring& password) {
    if (!m_archive) return false;

    // Find first non-directory file
    UInt32 numItems = 0;
    m_archive->GetNumberOfItems(&numItems);

    UInt32 testIndex = (UInt32)-1;
    for (UInt32 i = 0; i < numItems; i++) {
        PROPVARIANT prop;
        PropVariantInit(&prop);
        m_archive->GetProperty(i, kpidIsDir, &prop);
        bool isDir = (prop.vt == VT_BOOL && prop.boolVal != VARIANT_FALSE);
        PropVariantClear(&prop);

        if (!isDir) {
            testIndex = i;
            break;
        }
    }

    if (testIndex == (UInt32)-1) {
        return true;  // No files to test, assume success
    }

    // Create temp directory
    wchar_t tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempPath);
    std::wstring testDir = std::wstring(tempPath) + L"7ztest_" + std::to_wstring(GetTickCount64());
    CreateDirectoryW(testDir.c_str(), NULL);

    // Try to extract first file
    CExtractCallback* callback = new CExtractCallback(m_archive, testDir, password, nullptr);
    callback->AddRef();
    UInt32 indices[1] = { testIndex };
    HRESULT hr = m_archive->Extract(indices, 1, 0, callback);

    bool passwordWasRequested = callback->WasPasswordRequested();
    callback->Release();

    // Cleanup
    WIN32_FIND_DATAW fd;
    std::wstring searchPath = testDir + L"\\*";
    HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
    if (hFind != INVALID_HANDLE_VALUE) {
        do {
            if (wcscmp(fd.cFileName, L".") != 0 && wcscmp(fd.cFileName, L"..") != 0) {
                std::wstring filePath = testDir + L"\\" + fd.cFileName;
                DeleteFileW(filePath.c_str());
            }
        } while (FindNextFileW(hFind, &fd));
        FindClose(hFind);
    }
    RemoveDirectoryW(testDir.c_str());

    m_needsPassword = passwordWasRequested;

    return SUCCEEDED(hr);
}

bool SevenZipCore::Extract(const std::wstring& outDir,
                           const std::wstring& password,
                           ProgressCallback progress) {
    if (!m_archive) return false;

    CExtractCallback* callback = new CExtractCallback(m_archive, outDir, password, progress);
    callback->AddRef();
    HRESULT hr = m_archive->Extract(nullptr, (UInt32)-1, 0, callback);
    callback->Release();

    return SUCCEEDED(hr);
}

bool SevenZipCore::ExtractFiles(const std::vector<uint32_t>& indices,
                                const std::wstring& outDir,
                                const std::wstring& password,
                                ProgressCallback progress) {
    if (!m_archive || indices.empty()) return false;

    CExtractCallback* callback = new CExtractCallback(m_archive, outDir, password, progress);
    callback->AddRef();
    HRESULT hr = m_archive->Extract(indices.data(), (UInt32)indices.size(), 0, callback);
    callback->Release();

    return SUCCEEDED(hr);
}

bool SevenZipCore::Compress(const std::vector<std::wstring>& srcPaths,
                            const std::wstring& archivePath,
                            const std::wstring& format,
                            ProgressCallback progress) {
    // Find format
    const GUID* formatId = nullptr;
    for (const auto& fmt : m_formats) {
        if (_wcsicmp(format.c_str(), fmt.name.c_str()) == 0 ||
            _wcsicmp((L"." + format).c_str(), fmt.extension.c_str()) == 0) {
            if (fmt.canUpdate) {
                formatId = &fmt.classId;
                break;
            }
        }
    }

    if (!formatId) {
        // Default to 7z
        formatId = &CLSID_CFormat7z;
    }

    // Create output archive
    IOutArchive* outArchive = CreateOutArchive(*formatId);
    if (!outArchive) {
        return false;
    }

    // Create output stream
    CFullOutFileStream* outStream = new CFullOutFileStream();
    outStream->AddRef();
    if (!outStream->Create(archivePath.c_str())) {
        outStream->Release();
        outArchive->Release();
        return false;
    }

    // Create update callback
    CUpdateCallback* callback = new CUpdateCallback(srcPaths, progress);
    callback->AddRef();

    // Update archive
    HRESULT hr = outArchive->UpdateItems(outStream, callback->GetItemCount(), callback);

    callback->Release();
    outStream->Release();
    outArchive->Release();

    if (FAILED(hr)) {
        DeleteFileW(archivePath.c_str());
        return false;
    }

    return true;
}
