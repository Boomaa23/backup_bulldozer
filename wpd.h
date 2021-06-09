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

bool wpdInitialize();
void wpdUninitialize();

void GetClientInformation(IPortableDeviceValues** ppClientInformation);
void ChooseDevice(IPortableDevice** ppDevice);
void UnregisterForEventNotifications(_In_opt_ IPortableDevice *device, _In_opt_ PCWSTR eventCookie);

void RecursiveEnumerate(_In_ PCWSTR objectID, _In_ IPortableDeviceContent* content);
void EnumerateAllContent(_In_ IPortableDevice* device);

void DisplayFriendlyName(IPortableDeviceManager *pPortableDeviceManager, PCWSTR pPnPDeviceID);
void DisplayManufacturer(IPortableDeviceManager* pPortableDeviceManager, PCWSTR pPnPDeviceID);
DWORD EnumerateAllDevices();

HRESULT GetStringValue(IPortableDeviceProperties *pProperties,PCWSTR pszObjectID, REFPROPERTYKEY key,CAtlStringW &strStringValue);
HRESULT StreamCopy(IStream *pDestStream, IStream *pSourceStream, DWORD cbTransferSize, DWORD *pcbWritten);
void TransferContentFromDevice(IPortableDevice* pDevice);