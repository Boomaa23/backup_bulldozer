#pragma once

// General/main
#include <iostream>
#include <windows.h>
#include <vector>
#include <sstream>
#include <filesystem>
#include <string>
#include <thread>

// WPD/ATL
#include <PortableDeviceApi.h>
#include <PortableDevice.h>
#include <atlbase.h>
#include <atlstr.h>
#include <commdlg.h>
#include <atlbase.h>
#include <atlstr.h>
#include <atlcoll.h>
#include <specstrings.h>
#include <commdlg.h>
#include <new>
#include <strsafe.h>

struct Drive {
    std::string path;
    std::string name;
};

struct IndexedDrive : Drive {
    UINT index;
    bool isWPD;
};

struct WPDevice : Drive {
    PCWSTR deviceId;
};