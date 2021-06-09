#include "common.h"

bool wpdInitialize();
void wpdUninitialize();

void GetClientInformation(IPortableDeviceValues** ppClientInformation);
void ChooseDevice(IPortableDevice** ppDevice, UINT uiCurrentDevice, DWORD cPnPDeviceIDs);
void UnregisterForEventNotifications(_In_opt_ IPortableDevice *device, _In_opt_ PCWSTR eventCookie);

void RecursiveEnumerate(_In_ PCWSTR objectID, _In_ IPortableDeviceContent* content, std::vector<std::string>* deviceObjectIds);
void GetAllContent(_In_ IPortableDevice* device, std::vector<std::string>* deviceObjectIds);

std::string GetDeviceName(IPortableDeviceManager *pPortableDeviceManager, PCWSTR pPnPDeviceID);
std::vector<WPDevice> GetAllDevices();

HRESULT GetStringValue(IPortableDeviceProperties *pProperties,PCWSTR pszObjectID, REFPROPERTYKEY key,CAtlStringW &strStringValue);
HRESULT StreamCopy(IStream *pDestStream, IStream *pSourceStream, DWORD cbTransferSize, DWORD *pcbWritten);
void TransferContentFromDevice(IPortableDevice* pDevice);