// GuidInit.cpp - Initialize all 7-Zip GUIDs
// This file must be compiled once per project to instantiate the GUIDs
// DO NOT apply force-include to this file

// IMPORTANT: MyInitGuid.h must be included BEFORE any 7-Zip interface headers
// that use DEFINE_GUID/Z7_DEFINE_GUID macros

#include "Common/MyInitGuid.h"

// Include all interface headers that define GUIDs
#include "7zip/IDecl.h"
#include "7zip/ICoder.h"
#include "7zip/IStream.h"
#include "7zip/IProgress.h"
#include "7zip/IPassword.h"
#include "7zip/Archive/IArchive.h"

// Define CLSID_CArchiveHandler (same as in ArchiveExports.cpp)
Z7_DEFINE_GUID(CLSID_CArchiveHandler,
    k_7zip_GUID_Data1,
    k_7zip_GUID_Data2,
    k_7zip_GUID_Data3_Common,
    0x10, 0x00, 0x00, 0x01, 0x10, 0x00, 0x00, 0x00);

// CreateArchiver is defined in ArchiveExports.cpp
STDAPI CreateArchiver(const GUID *clsid, const GUID *iid, void **outObject);

// CreateObject wrapper (normally in DllExports.cpp but that has DllMain conflict)
STDAPI CreateObject(const GUID *clsid, const GUID *iid, void **outObject)
{
    return CreateArchiver(clsid, iid, outObject);
}
