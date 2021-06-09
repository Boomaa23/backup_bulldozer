#include "wpd.h"

// This number controls how many object identifiers are requested during each call
// to IEnumPortableDeviceObjectIDs::Next()
#define NUM_OBJECTS_TO_REQUEST  10

#define CLIENT_NAME             L"WPD Sample Application"
#define CLIENT_MAJOR_VER        1
#define CLIENT_MINOR_VER        0
#define CLIENT_REVISION         2

bool wpdInitialize() {
    // Enable the heap manager to terminate the process on heap error.
    (void)HeapSetInformation(NULL, HeapEnableTerminationOnCorruption, NULL, 0);

    // Initialize COM for COINIT_MULTITHREADED
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    return SUCCEEDED(hr);
}

void wpdUninitialize() {
    CoUninitialize();
}

// Creates and populates an IPortableDeviceValues with information about
// this application.  The IPortableDeviceValues is used as a parameter
// when calling the IPortableDevice::Open() method.
void GetClientInformation(IPortableDeviceValues** ppClientInformation) {
    // Client information is optional.  The client can choose to identify itself, or
    // to remain unknown to the driver.  It is beneficial to identify yourself because
    // drivers may be able to optimize their behavior for known clients. (e.g. An
    // IHV may want their bundled driver to perform differently when connected to their
    // bundled software.)

    // CoCreate an IPortableDeviceValues interface to hold the client information.
    HRESULT hr = CoCreateInstance(CLSID_PortableDeviceValues,
                                  nullptr,
                                  CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(ppClientInformation));
    if (SUCCEEDED(hr)) {
        // Attempt to set all bits of client information
        hr = (*ppClientInformation)->SetStringValue(WPD_CLIENT_NAME, CLIENT_NAME);
        if (FAILED(hr)) {
            printf("! Failed to set WPD_CLIENT_NAME, hr = 0x%lx\n",hr);
        }

        hr = (*ppClientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_MAJOR_VERSION, CLIENT_MAJOR_VER);
        if (FAILED(hr)) {
            printf("! Failed to set WPD_CLIENT_MAJOR_VERSION, hr = 0x%lx\n",hr);
        }

        hr = (*ppClientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_MINOR_VERSION, CLIENT_MINOR_VER);
        if (FAILED(hr)) {
            printf("! Failed to set WPD_CLIENT_MINOR_VERSION, hr = 0x%lx\n",hr);
        }

        hr = (*ppClientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_REVISION, CLIENT_REVISION);
        if (FAILED(hr)) {
            printf("! Failed to set WPD_CLIENT_REVISION, hr = 0x%lx\n",hr);
        }

        //  Some device drivers need to impersonate the caller in order to function correctly.  Since our application does not
        //  need to restrict its identity, specify SECURITY_IMPERSONATION so that we work with all devices.
        hr = (*ppClientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE, SECURITY_IMPERSONATION);
        if (FAILED(hr)) {
            printf("! Failed to set WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE, hr = 0x%lx\n",hr);
        }
    } else {
        printf("! Failed to CoCreateInstance CLSID_PortableDeviceValues, hr = 0x%lx\n",hr);
    }
}

// Calls EnumerateDevices() function to display devices on the system
// and to obtain the total number of devices found.  If 1 or more devices
// are found, this function prompts the user to choose a device using
// a zero-based index.
void ChooseDevice(IPortableDevice** ppDevice, const UINT uiCurrentDevice, DWORD cPnPDeviceIDs) {
    HRESULT                         hr              = S_OK;
    PWSTR*                          pPnpDeviceIDs   = NULL;
    CComPtr<IPortableDeviceManager> pPortableDeviceManager;
    CComPtr<IPortableDeviceValues>  pClientInformation;

    if (ppDevice == nullptr) {
        printf("! A nullptr IPortableDevice interface pointer was received\n");
        return;
    }

    if (*ppDevice != nullptr) {
        // To avoid operating on potiential bad pointers, reject any non-null
        // IPortableDevice interface pointers.  This will force the caller
        // to properly release the interface before obtaining a new one.
        printf("! A non-nullptr IPortableDevice interface pointer was received, please release this interface before obtaining a new one.\n");
        return;
    }

    // Fill out information about your application, so the device knows
    // who they are speaking to.

    GetClientInformation(&pClientInformation);

    // CoCreate the IPortableDeviceManager interface to enumerate
    // portable devices and to get information about them.
    hr = CoCreateInstance(CLSID_PortableDeviceManager,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_PPV_ARGS(&pPortableDeviceManager));
    if (FAILED(hr)) {
        printf("! Failed to CoCreateInstance CLSID_PortableDeviceManager, hr = 0x%lx\n",hr);
    }

    if (SUCCEEDED(hr)) {
        pPnpDeviceIDs = new (std::nothrow)PWSTR[cPnPDeviceIDs];
        hr = pPortableDeviceManager->GetDevices(pPnpDeviceIDs, &cPnPDeviceIDs);
        if (SUCCEEDED(hr)) {
            // CoCreate the IPortableDevice interface and call Open() with
        // the chosen PnPDeviceID string.
            hr = CoCreateInstance(CLSID_PortableDeviceFTM,
                nullptr,
                CLSCTX_INPROC_SERVER,
                IID_PPV_ARGS(ppDevice));
            if (SUCCEEDED(hr)) {
                // FORCE READ-ONLY, DO NOT WRITE
                //pClientInformation->SetUnsignedIntegerValue(WPD_CLIENT_DESIRED_ACCESS, GENERIC_READ);
                hr = (*ppDevice)->Open(pPnpDeviceIDs[uiCurrentDevice], pClientInformation);
                if (FAILED(hr)) {
                    printf("! Failed to Open the device, hr = 0x%lx\n", hr);
                    // Release the IPortableDevice interface, because we cannot proceed
                    // with an unopen device.
                    (*ppDevice)->Release();
                    *ppDevice = nullptr;
                }
            } else {
                printf("! Failed to CoCreateInstance CLSID_PortableDeviceFTM, hr = 0x%lx\n", hr);
            }
        } else {
            printf("! Failed to get the device list from the system, hr = 0x%lx\n", hr);
        }

        // Free all returned PnPDeviceID strings by using CoTaskMemFree.
        // NOTE: CoTaskMemFree can handle nullptr pointers, so no nullptr
        //       check is needed.
        for (DWORD dwIndex = 0; dwIndex < cPnPDeviceIDs; dwIndex++) {
            CoTaskMemFree(pPnpDeviceIDs[dwIndex]);
            pPnpDeviceIDs[dwIndex] = nullptr;
        }

        // Delete the array of PWSTR pointers
        delete[] pPnpDeviceIDs;
        pPnpDeviceIDs = nullptr;
        } else {
            printf("! Failed to get the device list from the system, hr = 0x%lx\n", hr);
        }
        // If no devices were found on the system, just exit this function.
}

void UnregisterForEventNotifications(_In_opt_ IPortableDevice *device, _In_opt_ PCWSTR eventCookie) {
    if (device == nullptr || eventCookie == nullptr) {
        return;
    }

    HRESULT hr = device->Unadvise(eventCookie);
    if (FAILED(hr)) {
        wprintf(L"! Failed to unregister for device events using registration cookie '%ws', hr = 0x%lx\n", eventCookie, hr);
    } else {
        wprintf(L"This application used the registration cookie '%ws' to unregister from receiving device event notifications", eventCookie);
    }
}

// Recursively called function which enumerates using the specified
// object identifier as the parent.
void RecursiveEnumerate(_In_ PCWSTR objectID, _In_ IPortableDeviceContent* content, _In_ std::vector<std::string>* deviceObjectIds) {
    CComPtr<IEnumPortableDeviceObjectIDs> enumObjectIDs;

    // Get an IEnumPortableDeviceObjectIDs interface by calling EnumObjects with the
    // specified parent object identifier.
    HRESULT hr = content->EnumObjects(0,                // Flags are unused
                                      objectID,         // Starting from the passed in object
                                      nullptr,          // Filter is unused
                                      &enumObjectIDs);
    if (FAILED(hr)) {
        wprintf(L"! Failed to get IEnumPortableDeviceObjectIDs from IPortableDeviceContent, hr = 0x%lx\n", hr);
    }

    // Loop calling Next() while S_OK is being returned.
    while (hr == S_OK) {
        DWORD  numFetched = 0;
        PWSTR  objectIDArray[NUM_OBJECTS_TO_REQUEST] = {nullptr};
        hr = enumObjectIDs->Next(NUM_OBJECTS_TO_REQUEST,    // Number of objects to request on each NEXT call
                                 objectIDArray,             // Array of PWSTR array which will be populated on each NEXT call
                                 &numFetched);              // Number of objects written to the PWSTR array
        if (SUCCEEDED(hr)) {
            // Traverse the results of the Next() operation and recursively enumerate
            // Remember to free all returned object identifiers using CoTaskMemFree()
            
            for (DWORD index = 0; (index < numFetched) && (objectIDArray[index] != nullptr); index++) {
                RecursiveEnumerate(objectIDArray[index], content, deviceObjectIds);
                char tempId[8];
                wcstombs(tempId, objectIDArray[index], 8);
                deviceObjectIds->push_back(tempId);
                // Free allocated PWSTRs after the recursive enumeration call has completed.
                CoTaskMemFree(objectIDArray[index]);
                objectIDArray[index] = nullptr;
            }
        }
    }
}

// Enumerate all content on the device starting with the
// "DEVICE" object
void GetAllContent(_In_ IPortableDevice* device, std::vector<std::string>* deviceObjectIds) {
    HRESULT                         hr = S_OK;
    CComPtr<IPortableDeviceContent>  content;

    // Get an IPortableDeviceContent interface from the IPortableDevice interface to
    // access the content-specific methods.
    hr = device->Content(&content);
    if (FAILED(hr)) {
        wprintf(L"! Failed to get IPortableDeviceContent from IPortableDevice, hr = 0x%lx\n", hr);
    }

    // Enumerate content starting from the "DEVICE" object.
    if (SUCCEEDED(hr)) {
        wprintf(L"\n");
        RecursiveEnumerate(WPD_DEVICE_OBJECT_ID, content.Detach(), deviceObjectIds);
    }
}

// Reads and displays the device friendly name for the specified PnPDeviceID string
std::string GetDeviceName(IPortableDeviceManager *pPortableDeviceManager, PCWSTR pPnPDeviceID) {
    DWORD   cchFriendlyName = 0;
    wchar_t*   pszFriendlyName = nullptr;

    // First, pass nullptr as the PWSTR return string parameter to get the total number
    // of characters to allocate for the string value.
    HRESULT hr = pPortableDeviceManager->GetDeviceFriendlyName(pPnPDeviceID, nullptr, &cchFriendlyName);
    if (FAILED(hr)) {
        printf("! Failed to get number of characters for device friendly name, hr = 0x%lx\n",hr);
    }

    // Second allocate the number of characters needed and retrieve the string value.
    if ((hr == S_OK) && (cchFriendlyName > 0)) {
        pszFriendlyName = new (std::nothrow) WCHAR[cchFriendlyName];
        if (pszFriendlyName != nullptr) {
            hr = pPortableDeviceManager->GetDeviceFriendlyName(pPnPDeviceID, pszFriendlyName, &cchFriendlyName);
            if (SUCCEEDED(hr)) {
                char outName[256];
                wcstombs(outName, pszFriendlyName, 256);
                return outName;
            } else {
                printf("! Failed to get device friendly name, hr = 0x%lx\n",hr);
            }

            // Delete the allocated friendly name string
            delete [] pszFriendlyName;
            pszFriendlyName = nullptr;
        } else {
            printf("! Failed to allocate memory for the device friendly name string\n");
        }
    }

    if (SUCCEEDED(hr) && (cchFriendlyName == 0)) {
        printf("The device did not provide a friendly name.\n");
    }
}

// Enumerates all Windows Portable Devices, displays the friendly name,
// manufacturer, and description of each device.  This function also
// returns the total number of devices found.
std::vector<WPDevice> GetAllDevices() {
    std::vector<WPDevice>           devices;
    DWORD                           cPnPDeviceIDs = 0;
    PWSTR*                          pPnpDeviceIDs = nullptr;
    CComPtr<IPortableDeviceManager> pPortableDeviceManager;

    // CoCreate the IPortableDeviceManager interface to enumerate
    // portable devices and to get information about them.
    HRESULT hr = CoCreateInstance(CLSID_PortableDeviceManager,
                                  nullptr,
                                  CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(&pPortableDeviceManager));
    if (FAILED(hr)) {
        printf("! Failed to CoCreateInstance CLSID_PortableDeviceManager, hr = 0x%lx\n",hr);
    }

    // First, pass nullptr as the PWSTR array pointer to get the total number
    // of devices found on the system.
    if (SUCCEEDED(hr)) {
        hr = pPortableDeviceManager->GetDevices(nullptr, &cPnPDeviceIDs);
        if (FAILED(hr)) {
            printf("! Failed to get number of devices on the system, hr = 0x%lx\n",hr);
        }
    }

    // Report the number of devices found.  NOTE: we will report 0, if an error
    // occured.

    // Second, allocate an array to hold the PnPDeviceID strings returned from
    // the IPortableDeviceManager::GetDevices method
    if (SUCCEEDED(hr) && (cPnPDeviceIDs > 0)) {
        pPnpDeviceIDs = new (std::nothrow) PWSTR[cPnPDeviceIDs];
        if (pPnpDeviceIDs != nullptr) {
            DWORD dwIndex = 0;

            hr = pPortableDeviceManager->GetDevices(pPnpDeviceIDs, &cPnPDeviceIDs);
            if (SUCCEEDED(hr)) {
                // For each device found, display the devices friendly name,
                // manufacturer, and description strings.
                for (dwIndex = 0; dwIndex < cPnPDeviceIDs; dwIndex++) {
                    devices.push_back(
                        WPDevice{
                            "WPD",
                            GetDeviceName(pPortableDeviceManager, pPnpDeviceIDs[dwIndex]), 
                            pPnpDeviceIDs[dwIndex]
                        }
                    );
                }
            } else {
                printf("! Failed to get the device list from the system, hr = 0x%lx\n",hr);
            }

            // Free all returned PnPDeviceID strings by using CoTaskMemFree.
            // NOTE: CoTaskMemFree can handle nullptr pointers, so no nullptr
            //       check is needed.
            for (dwIndex = 0; dwIndex < cPnPDeviceIDs; dwIndex++) {
                CoTaskMemFree(pPnpDeviceIDs[dwIndex]);
                pPnpDeviceIDs[dwIndex] = nullptr;
            }

            // Delete the array of PWSTR pointers
            delete [] pPnpDeviceIDs;
            pPnpDeviceIDs = nullptr;
        } else {
            printf("! Failed to allocate memory for PWSTR array\n");
        }
    }
    return devices;
}

// Reads a string property from the IPortableDeviceProperties
// interface and returns it in the form of a CAtlStringW
HRESULT GetStringValue(IPortableDeviceProperties *pProperties,PCWSTR pszObjectID, 
        REFPROPERTYKEY key,CAtlStringW &strStringValue) {
    CComPtr<IPortableDeviceValues>        pObjectProperties;
    CComPtr<IPortableDeviceKeyCollection> pPropertiesToRead;

    // 1) CoCreate an IPortableDeviceKeyCollection interface to hold the the property key
    // we wish to read.
    HRESULT hr = CoCreateInstance(CLSID_PortableDeviceKeyCollection,
                                  nullptr,
                                  CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(&pPropertiesToRead));

    // 2) Populate the IPortableDeviceKeyCollection with the keys we wish to read.
    // NOTE: We are not handling any special error cases here so we can proceed with
    // adding as many of the target properties as we can.
    if (SUCCEEDED(hr)) {
        if (pPropertiesToRead != nullptr) {
            HRESULT hrTemp = S_OK;
            hrTemp = pPropertiesToRead->Add(key);
            if (FAILED(hrTemp)) {
                printf("! Failed to add PROPERTYKEY to IPortableDeviceKeyCollection, hr= 0x%lx\n", hrTemp);
            }
        }
    }

    // 3) Call GetValues() passing the collection of specified PROPERTYKEYs.
    if (SUCCEEDED(hr)) {
        hr = pProperties->GetValues(pszObjectID,         // The object whose properties we are reading
                                    pPropertiesToRead,   // The properties we want to read
                                    &pObjectProperties); // Driver supplied property values for the specified object

        // The error is handled by the caller, which also displays an error message to the user.
    }

    // 4) Extract the string value from the returned property collection
    if (SUCCEEDED(hr)) {
        PWSTR pszStringValue = nullptr;
        hr = pObjectProperties->GetStringValue(key, &pszStringValue);
        if (SUCCEEDED(hr)) {
            // Assign the newly read string to the CAtlStringW out
            // parameter.
            strStringValue = pszStringValue;
        } else {
            printf("! Failed to find property in IPortableDeviceValues, hr = 0x%lx\n",hr);
        }

        CoTaskMemFree(pszStringValue);
        pszStringValue = nullptr;
    }

    return hr;
}

// Copies data from a source stream to a destination stream using the
// specified cbTransferSize as the temporary buffer size.
HRESULT StreamCopy(IStream *pDestStream, IStream *pSourceStream, DWORD cbTransferSize, DWORD *pcbWritten) {
    HRESULT hr = S_OK;

    // Allocate a temporary buffer (of Optimal transfer size) for the read results to
    // be written to.
    BYTE *pObjectData = new (std::nothrow) BYTE[cbTransferSize];
    if (pObjectData != nullptr) {
        DWORD cbTotalBytesRead    = 0;
        DWORD cbTotalBytesWritten = 0;

        DWORD cbBytesRead    = 0;
        DWORD cbBytesWritten = 0;

        // Read until the number of bytes returned from the source stream is 0, or
        // an error occured during transfer.
        do {
            // Read object data from the source stream
            hr = pSourceStream->Read(pObjectData, cbTransferSize, &cbBytesRead);
            if (FAILED(hr)) {
                printf("! Failed to read %lu bytes from the source stream, hr = 0x%lx\n",cbTransferSize, hr);
            }

            // Write object data to the destination stream
            if (SUCCEEDED(hr)) {
                cbTotalBytesRead += cbBytesRead; // Calculating total bytes read from device for debugging purposes only

                hr = pDestStream->Write(pObjectData, cbBytesRead, &cbBytesWritten);
                if (FAILED(hr)) {
                    printf("! Failed to write %lu bytes of object data to the destination stream, hr = 0x%lx\n",cbBytesRead, hr);
                }

                if (SUCCEEDED(hr)) {
                    cbTotalBytesWritten += cbBytesWritten; // Calculating total bytes written to the file for debugging purposes only
                }
            }

            // Output Read/Write operation information only if we have received data and if no error has occured so far.
            if (SUCCEEDED(hr) && (cbBytesRead > 0)) {
                printf("Read %lu bytes from the source stream...Wrote %lu bytes to the destination stream...\n", cbBytesRead, cbBytesWritten);
            }

        } while (SUCCEEDED(hr) && (cbBytesRead > 0));

        // If the caller supplied a pcbWritten parameter and we
        // and we are successful, set it to cbTotalBytesWritten
        // before exiting.
        if ((SUCCEEDED(hr)) && (pcbWritten != nullptr)) {
            *pcbWritten = cbTotalBytesWritten;
        }

        // Remember to delete the temporary transfer buffer
        delete [] pObjectData;
        pObjectData = nullptr;
    } else {
        printf("! Failed to allocate %lu bytes for the temporary transfer buffer.\n", cbTransferSize);
    }

    return hr;
}

void TransferContentFromDevice(IPortableDevice* pDevice) {
    HRESULT                            hr                   = S_OK;
    WCHAR                              szSelection[81]      = {0};
    CComPtr<IPortableDeviceContent>    pContent;
    CComPtr<IPortableDeviceResources>  pResources;
    CComPtr<IPortableDeviceProperties> pProperties;
    CComPtr<IStream>                   pObjectDataStream;
    CComPtr<IStream>                   pFinalFileStream;
    DWORD                              cbOptimalTransferSize = 0;
    CAtlStringW                        strOriginalFileName;

    if (pDevice == nullptr) {
        printf("! A nullptr IPortableDevice interface pointer was received\n");
        return;
    }

    // Prompt user to enter an object identifier on the device to transfer.
    printf("Enter the identifier of the object you wish to transfer.\n>");
    hr = StringCbGetsW(szSelection,sizeof(szSelection));
    if (FAILED(hr)) {
        printf("An invalid object identifier was specified, aborting content transfer\n");
    }
    // 1) get an IPortableDeviceContent interface from the IPortableDevice interface to
    // access the content-specific methods.
    if (SUCCEEDED(hr)) {
        hr = pDevice->Content(&pContent);
        if (FAILED(hr)) {
            printf("! Failed to get IPortableDeviceContent from IPortableDevice, hr = 0x%lx\n",hr);
        }
    }
    // 2) Get an IPortableDeviceResources interface from the IPortableDeviceContent interface to
    // access the resource-specific methods.
    if (SUCCEEDED(hr)) {
        hr = pContent->Transfer(&pResources);
        if (FAILED(hr)) {
            printf("! Failed to get IPortableDeviceResources from IPortableDeviceContent, hr = 0x%lx\n",hr);
        }
    }
    // 3) Get the IStream (with READ access) and the optimal transfer buffer size
    // to begin the transfer.
    if (SUCCEEDED(hr)) {
        hr = pResources->GetStream(szSelection,             // Identifier of the object we want to transfer
                                   WPD_RESOURCE_DEFAULT,    // We are transferring the default resource (which is the entire object's data)
                                   STGM_READ,               // Opening a stream in READ mode, because we are reading data from the device.
                                   &cbOptimalTransferSize,  // Driver supplied optimal transfer size
                                   &pObjectDataStream);
        if (FAILED(hr)) {
            printf("! Failed to get IStream (representing object data on the device) from IPortableDeviceResources, hr = 0x%lx\n",hr);
        }
    }

    // 4) Read the WPD_OBJECT_ORIGINAL_FILE_NAME property so we can properly name the
    // transferred object.  Some content objects may not have this property, so a
    // fall-back case has been provided below. (i.e. Creating a file named <objectID>.data )
    if (SUCCEEDED(hr)) {
        hr = pContent->Properties(&pProperties);
        if (SUCCEEDED(hr)) {
            hr = GetStringValue(pProperties,
                                szSelection,
                                WPD_OBJECT_ORIGINAL_FILE_NAME,
                                strOriginalFileName);
            if (FAILED(hr)) {
                printf("! Failed to read WPD_OBJECT_ORIGINAL_FILE_NAME on object '%ws', hr = 0x%lx\n", szSelection, hr);
                strOriginalFileName.Format(L"%ws.data", szSelection);
                printf("* Creating a filename '%ws' as a default.\n", (PWSTR)strOriginalFileName.GetString());
                // Set the HRESULT to S_OK, so we can continue with our newly generated
                // temporary file name.
                hr = S_OK;
            }
        } else {
            printf("! Failed to get IPortableDeviceProperties from IPortableDeviceContent, hr = 0x%lx\n", hr);
        }
    }
    // 5) Create a destination for the data to be written to.  In this example we are
    // creating a temporary file which is named the same as the object identifier string.
    if (SUCCEEDED(hr)) {
        hr = SHCreateStreamOnFile((LPCSTR)strOriginalFileName.GetString(), STGM_CREATE | STGM_WRITE, &pFinalFileStream);
        if (FAILED(hr)) {
            printf("! Failed to create a temporary file named (%ws) to transfer object (%ws), hr = 0x%lx\n",(PWSTR)strOriginalFileName.GetString(), szSelection, hr);
        }
    }
    // 6) Read on the object's data stream and write to the final file's data stream using the
    // driver supplied optimal transfer buffer size.
    if (SUCCEEDED(hr)) {
        DWORD cbTotalBytesWritten = 0;

        // Since we have IStream-compatible interfaces, call our helper function
        // that copies the contents of a source stream into a destination stream.
        hr = StreamCopy(pFinalFileStream,       // Destination (The Final File to transfer to)
                        pObjectDataStream,      // Source (The Object's data to transfer from)
                        cbOptimalTransferSize,  // The driver specified optimal transfer buffer size
                        &cbTotalBytesWritten);  // The total number of bytes transferred from device to the finished file
        if (FAILED(hr)) {
            printf("! Failed to transfer object from device, hr = 0x%lx\n",hr);
        } else {
            printf("* Transferred object '%ws' to '%ws'.\n", szSelection, (PWSTR)strOriginalFileName.GetString());
        }
    }
}