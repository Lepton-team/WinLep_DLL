// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include <objbase.h>
#include <shlwapi.h>
#include <thumbcache.h> // For IThumbnailProvider.
#include <shlobj.h>     // For SHChangeNotify
#include <new>
#include <combaseapi.h>
#include <Windows.h>
#include <strsafe.h>

#include "dllmain.h"

// Generated UUID: d0385023-d398-4b09-89a2-aa1ade3cdfe7
#define SZ_CLSID_LEPTON_THUMBNAIL_HANDLER L"{d0385023-d398-4b09-89a2-aa1ade3cdfe7}"
//#define SZ_LEPTON_THUMBNAIL_HANDLER = L"Lepton Thumbnail Handler"
#define THUMBNAIL_IMAGE_HANDLER_SUBKEY L"{E357FCCD-A995-4576-B01F-234630154E96}"

unsigned long dllRefCount = 0;

const CLSID CLSID_LeptonThumbnailProvider = { 0xd0385023, 0xd398, 0x4b09, {0x89, 0xa2, 0xaa, 0x1a, 0xde, 0x3c, 0xdf, 0xe7} };

// hInstance = "handle to an instance"
// The operating system uses this value to identify the executable when it is loaded in memory.
// The instance handle is needed to load icons or bitmaps.
HINSTANCE hInstance = nullptr;

typedef struct REGKEY_DELETEKEY {
	HKEY hKey;
	LPCWSTR lpszSubKey;
} REGKEY_DELETEKEY;

typedef struct REGKEY_SUBKEY_AND_VALUE {
	HKEY hKey;
	LPCWSTR lpszSubKey;
	LPCWSTR lpszValue;
	DWORD dwType;
	DWORD_PTR dwData;
} REGKEY_SUBKEY_AND_VALUE;

HRESULT createRegistryKeys(REGKEY_SUBKEY_AND_VALUE*, ULONG);
HRESULT createRegistryKey(REGKEY_SUBKEY_AND_VALUE*);

HRESULT deleteRegistryKeys(REGKEY_DELETEKEY*, ULONG);

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved) {
	switch (ul_reason_for_call) {
	case DLL_PROCESS_ATTACH:
		hInstance = hModule;
		DisableThreadLibraryCalls(hModule);
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}

	return TRUE;
}

STDAPI DllRegisterServer() {
	HRESULT result;
	WCHAR szModule[MAX_PATH] = { 0 };

	if (GetModuleFileName(hInstance, szModule, ARRAYSIZE(szModule) == 0)) {
		return HRESULT_FROM_WIN32(GetLastError());
	}

	//Register lepton thumbnail provider
	//TODO: file extsion name ... For now it's .lep // FIXME
	REGKEY_SUBKEY_AND_VALUE thumbnailProviderKeys[] = {
		{HKEY_CLASSES_ROOT, L"CLSID\\" SZ_CLSID_LEPTON_THUMBNAIL_HANDLER, nullptr, REG_SZ, (DWORD_PTR)L"Lepton Thumbnail Provider"},
		{HKEY_CLASSES_ROOT, L"CLSID\\" SZ_CLSID_LEPTON_THUMBNAIL_HANDLER L"\\InprocServer32", L"ThreadingModel", REG_SZ, (DWORD_PTR)L"Apartment"},
		{HKEY_CLASSES_ROOT, L".lep\\shellex\\" THUMBNAIL_IMAGE_HANDLER_SUBKEY, nullptr, REG_SZ, (DWORD_PTR)SZ_CLSID_LEPTON_THUMBNAIL_HANDLER}
	};

	result = createRegistryKeys(thumbnailProviderKeys, ARRAYSIZE(thumbnailProviderKeys));

	if (FAILED(result)) {
		REGKEY_DELETEKEY keysToDelete[] = {
			{HKEY_CLASSES_ROOT, L".lep\\shellex\\" THUMBNAIL_IMAGE_HANDLER_SUBKEY},
			{HKEY_CLASSES_ROOT, L"CLSID\\" SZ_CLSID_LEPTON_THUMBNAIL_HANDLER}
		};

		deleteRegistryKeys(keysToDelete, ARRAYSIZE(keysToDelete));
		return result;
	}

	//TODO: Register property handler

	// Notify shell about association changes
	SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);

	return S_OK;
}

STDAPI DllUnregisterServer() {
	WCHAR szModule[MAX_PATH] = { 0 };

	if (GetModuleFileName(hInstance, szModule, ARRAYSIZE(szModule)) == 0) {
		return HRESULT_FROM_WIN32(GetLastError());
	}

	REGKEY_DELETEKEY keysToDelete[] = {
		{HKEY_CLASSES_ROOT, L"CLSID\\" SZ_CLSID_LEPTON_THUMBNAIL_HANDLER},
		//{} //TODO: Property handler
	};

	HRESULT result = deleteRegistryKeys(keysToDelete, ARRAYSIZE(keysToDelete));

	// TODO: Filename extension!
	REGKEY_DELETEKEY keys2[] = {
		{ HKEY_CLASSES_ROOT, L".lep\\shellex\\" THUMBNAIL_IMAGE_HANDLER_SUBKEY},
		//{ HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\.kra" },
	};
	deleteRegistryKeys(keys2, ARRAYSIZE(keys2));

	return result;
}

STDAPI DllCanUnloadNow() {
	return dllRefCount > 0 ? S_FALSE : S_OK;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {
	return S_OK;
}

HRESULT deleteRegistryKeys(REGKEY_DELETEKEY* keys, ULONG size) {
	HRESULT res = S_OK;
	LSTATUS status;

	for (ULONG i = 0; i < size; i++) {
		status = RegDeleteTree(keys[i].hKey, keys[i].lpszSubKey);
		if (status != NOERROR) {
			res = HRESULT_FROM_WIN32(status);
		}
	}

	return res;
}

HRESULT createRegistryKeys(REGKEY_SUBKEY_AND_VALUE* keys, ULONG size) {
	HRESULT res = S_OK;
	for (ULONG i = 0; i < size; i++) {
		HRESULT tmp = createRegistryKey(&keys[i]);
		if (FAILED(tmp)) {
			res = tmp;
		}
	}

	return res;
}

HRESULT createRegistryKey(REGKEY_SUBKEY_AND_VALUE* key) {
	HRESULT res = S_OK;

	size_t dataSize;
	LPVOID data = nullptr;

	switch (key->dwType) {
	case REG_DWORD:
		data = (LPVOID)(LPDWORD)&key->dwData;
		dataSize = sizeof(DWORD);
		break;

	case REG_SZ:
	case REG_EXPAND_SZ:
		// Safe replacement for strlen
		res = StringCbLength((LPCWSTR)key->dwData, STRSAFE_MAX_CCH, &dataSize);
		if (SUCCEEDED(res)) {
			data = (LPVOID)(LPCWSTR)key->dwData;
			dataSize += sizeof(WCHAR);
		}
		break;

	default:
		res = E_INVALIDARG;
	}

	if (SUCCEEDED(res)) {
		LSTATUS status = SHSetValue(key->hKey, key->lpszSubKey, key->lpszValue, key->dwType, data, dataSize);
		if (status != NOERROR) {
			res = HRESULT_FROM_WIN32(status);
		}
	}

	return res;
}

void lepton::IncDllRef() {
	// Increments the value of the specified 32-bit variable as an atomic operation.
	InterlockedIncrement(&dllRefCount);
}

void lepton::DecDllRef() {
	// Decrements the value of the specified 32-bit variable as an atomic operation.
	InterlockedDecrement(&dllRefCount);
}