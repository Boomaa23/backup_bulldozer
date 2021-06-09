#include "common.h"
#include "wpd.h"

static const std::filesystem::directory_options FS_DIR_OPTS = (
        std::filesystem::directory_options::follow_directory_symlink |
        std::filesystem::directory_options::skip_permission_denied
);

static const std::string MONTHS[] = {
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December"
};

static const char BYTE_PREFIXES[] = {'k', 'M', 'G', 'T', 'P', 'E'};

struct CopyResult {
    bool success;
    LARGE_INTEGER sizeBytes;
};


std::string lastErrorMessage() {
    DWORD errMsgId = GetLastError();
    LPSTR msgBuf = nullptr;
    size_t msgSize = FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                                    nullptr, errMsgId, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                    (LPSTR) &msgBuf, 0, nullptr);
    std::string message(msgBuf, msgSize);
    LocalFree(msgBuf);
    return message;
}

int getCurrentMsTime() {
    return std::chrono::duration_cast<std::chrono::milliseconds>(
            std::chrono::system_clock::now().time_since_epoch()).count();
}

std::string printDrive(Drive *drive, bool toStdOut) {
    std::stringstream ss;
    ss << drive->path << " - " << drive->name << std::endl;
    if (toStdOut) {
        std::cout << ss.str();
    }
    return ss.str();
}

SYSTEMTIME getFileTime(HANDLE *file) {
    FILETIME fileTime;
    GetFileTime(*file, &fileTime, nullptr, nullptr);
    SYSTEMTIME sysTime;
    FileTimeToSystemTime(&fileTime, &sysTime);
    return sysTime;
}

std::vector<Drive> getLogicalDrives() {
    DWORD reqBufSize = GetLogicalDriveStringsA(0, nullptr);
    LPSTR driveLetters = new TCHAR[reqBufSize];
    GetLogicalDriveStringsA(reqBufSize, driveLetters);
    LPSTR loopDrive = driveLetters;
    std::vector<Drive> drives = {};
    do {
        if (strcmp(loopDrive, "C:\\") != 0) {
            LPSTR volName = new TCHAR[256];
            GetVolumeInformationA(loopDrive, volName, (DWORD) 256,
                    nullptr, nullptr, nullptr, nullptr, 0);
            drives.push_back(Drive{std::string(1, loopDrive[0]) + ":\\", volName});
            delete[] volName;
        }
        while (*loopDrive++);
    } while (*loopDrive);
    return drives;
}

CopyResult copyFile(std::string *srcPath, const Drive *srcDrive, const std::string *baseDstPath) {
    LPCSTR lpcSrcPath = srcPath->c_str();
    HANDLE srcFile = CreateFileA(lpcSrcPath, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    SYSTEMTIME time = getFileTime(&srcFile);

    std::stringstream dstPathSS;
    std::string filename = srcPath->substr(srcPath->find_last_of('\\') + 1);
    dstPathSS << *baseDstPath << '\\' << srcDrive->name << '\\' << time.wYear << '\\'
               << MONTHS[time.wMonth - 1] << '\\' << filename;

    std::string fileDstPath = dstPathSS.str();
    std::string dirDstPath = fileDstPath.substr(0, fileDstPath.find_last_of('\\'));

    size_t ctr = 0;
    do {
        ctr = dirDstPath.find_first_of("\\/", ctr + 1);
        CreateDirectory(dirDstPath.substr(0, ctr).c_str(), nullptr);
    } while (ctr != std::string::npos);

    bool success = CopyFileA(lpcSrcPath, fileDstPath.c_str(), TRUE);
    LARGE_INTEGER fileSize;
    GetFileSizeEx(srcFile, &fileSize);
    return CopyResult{success, fileSize};
}

IndexedDrive selectDrive(std::vector<Drive> *drives, std::vector<WPDevice> *wpDevices) {
    size_t i;
    for (i = 0; i < drives->size(); i++) {
        std::cout << i << ") " << printDrive(&drives->at(i), false);
    }
    for (i = 0; i < wpDevices->size(); i++) {
        std::cout << drives->size() + i << ") " << printDrive(&wpDevices->at(i), false);
    }
    std::string sel;
    do {
        std::getline(std::cin, sel);
    } while (sel.empty());
    UINT idx = std::stoi(sel);
    Drive *selDrive;
    bool isWPD;
    if (isWPD = i >= drives->size()) {
        idx -= drives->size();
        selDrive = &wpDevices->at(idx);
    } else {
        selDrive = &drives->at(idx);
    }
    return IndexedDrive{selDrive->path, selDrive->name, idx, isWPD};
}

std::string userInput(const std::string &preMessage, bool allowBlank) {
    std::string sel;
    do {
        std::cout << preMessage << std::endl;
        std::getline(std::cin, sel);
    } while (sel.empty() && !allowBlank);
    return sel;
}

std::string bytesHumanReadable(int64_t numBytes) {
    if (-1000 < numBytes && numBytes < 1000) {
        return std::to_string(numBytes) + " B";
    }
    size_t idx = 0;
    while (numBytes <= -999950 || numBytes >= 999950) {
        numBytes /= 1000;
        idx++;
    }
    char buf[32];
    snprintf(buf, sizeof(buf), "%.2f %cB", numBytes / 1000.0, BYTE_PREFIXES[idx]);
    return std::string(buf);
}

void cleanExtension(std::string *ext) {
    ext->erase(std::remove(ext->begin(), ext->end(), '.'), ext->end());
    transform(ext->begin(), ext->end(), ext->begin(), ::tolower);
}

int main() {
    wpdInitialize();
    std::vector<Drive> drives = getLogicalDrives();
    std::string out = userInput("Base destination path:", false);
    if (out.at(out.size() - 1) != '\\') {
        out.push_back('\\');
    }
    std::string fileType = userInput("File type (blank if all):", true);
    if (!fileType.empty()) {
        cleanExtension(&fileType);
    }

    CComPtr<IPortableDevice> portableDevice;
    std::vector<WPDevice> wpDevices = GetAllDevices();

    IndexedDrive selDrive = selectDrive(&drives, &wpDevices);
    if (selDrive.isWPD) {
        ChooseDevice(&portableDevice, selDrive.index, wpDevices.size());
        std::vector<std::string> deviceObjectIds;
        GetAllContent(portableDevice, &deviceObjectIds);
        return 0;
    }
    if (selDrive.name.empty()) {
        selDrive.name = userInput("Drive name missing, input new name:", false);
    }

    int64_t copiedBytes = 0;
    int64_t skippedBytes = 0;
    int startTime = getCurrentMsTime();
    std::cout << "Starting copy..." << std::endl;
    for (const std::filesystem::directory_entry &entry :
            std::filesystem::recursive_directory_iterator(selDrive.path, FS_DIR_OPTS)) {
        std::string inPath = entry.path().string();
        size_t extPos = inPath.find_last_of('.');
        bool skipType = fileType.empty();

        if (!skipType && extPos != std::string::npos) {
            std::string pathExt = inPath.substr(extPos);
            cleanExtension(&pathExt);
            skipType = strcmp(pathExt.c_str(), fileType.c_str()) == 0;
        }

        if (!entry.is_directory() && skipType) {
            CopyResult result = copyFile(&inPath, &selDrive, &out);
            if (result.success) {
                copiedBytes += result.sizeBytes.QuadPart;
            } else {
                skippedBytes += result.sizeBytes.QuadPart;
            }
        }
    }

    double elapsedTime = ((uint64_t)getCurrentMsTime() - startTime) / 1000.0;
    std::cout << bytesHumanReadable(copiedBytes)
        << " copied in " << elapsedTime << " seconds ("
        << bytesHumanReadable(copiedBytes / elapsedTime)
        << "/s) with " << bytesHumanReadable(skippedBytes) << " skipped." << std::endl;
    return 0;
}

