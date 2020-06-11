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
#define SZ_CLSID_WINLEP_THUMBNAIL_HANDLER L"{d0385023-d398-4b09-89a2-aa1ade3cdfe7}"
#define SZ_WINLEP_THUMBNAIL_HANDLER L"WinLep Thumbnail Handler"
#define SZ_THUMBNAIL_IMAGE_HANDLER_SUBKEY L"{E357FCCD-A995-4576-B01F-234630154E96}"

//#define SZ_CLSID_LEPTON_PROPERTY_HANDLER L"{b1e75cb3-ad2a-40ed-a836-f06f3f32a8b9}"
//#define SZ_LEPTON_PROPERTY_HANDLER L"Lepton Property Handler"

//TODO: Change file extsion name ?
#define SZ_FILE_EXTENSION L".wlep"

unsigned long dllRefCount = 0;

const CLSID CLSID_LeptonThumbnailProvider = {0xd0385023, 0xd398, 0x4b09, {0x89, 0xa2, 0xaa, 0x1a, 0xde, 0x3c, 0xdf, 0xe7}};
//const CLSID CLSID_LeptonPropertyHandler = {0xb1e75cb3, 0xad2a, 0x40ed, {0xa8, 0x36, 0xf0, 0x6f, 0x3f, 0x32, 0xa8, 0xb9}};

// hInstance = "handle to an instance"
// The operating system uses this value to identify the executable when it is loaded in memory.
// The instance handle is needed to load icons or bitmaps.
HINSTANCE g_hInstance = nullptr;

struct REGISTRY_ENTRY {
	HKEY hkeyRoot;
	PCWSTR pszKeyName;
	PCWSTR pszValueName;
	PCWSTR pszData;
};

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved) {
	if (ul_reason_for_call == DLL_PROCESS_ATTACH) {
		g_hInstance = hModule;
		DisableThreadLibraryCalls(hModule);
	}

	return TRUE;
}

typedef HRESULT(*PFNCREATEINSTANCE)(REFIID riid, void **ppvObject);

extern HRESULT WinLepThumbnailProvider_CreateInstance(REFIID riid, void **ppv);

struct CLASS_OBJECT_INIT {
	const CLSID *pClsid;
	PFNCREATEINSTANCE pfnCreate;
};

const CLASS_OBJECT_INIT c_rgClassObjectInit[] = {
	{ &CLSID_LeptonThumbnailProvider, WinLepThumbnailProvider_CreateInstance }
};

class CClassFactory : public IClassFactory {
public:
	static HRESULT CreateInstance(REFCLSID clsid, const CLASS_OBJECT_INIT *pClassObjectInits, size_t cClassObjectInits, REFIID riid, void **ppv) {
		*ppv = NULL;
		HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;
		for (size_t i = 0; i < cClassObjectInits; i++) {
			if (clsid == *pClassObjectInits[i].pClsid) {
				IClassFactory *pClassFactory = new (std::nothrow) CClassFactory(pClassObjectInits[i].pfnCreate);
				hr = pClassFactory ? S_OK : E_OUTOFMEMORY;
				if (SUCCEEDED(hr)) {
					hr = pClassFactory->QueryInterface(riid, ppv);
					pClassFactory->Release();
				}
				break; // match found
			}
		}
		return hr;
	}

	CClassFactory(PFNCREATEINSTANCE pfnCreate) : _cRef(1), _pfnCreate(pfnCreate) {
		wlep::IncDllRef();
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv) {
		static const QITAB qit[] = {
			QITABENT(CClassFactory, IClassFactory),
			{0}
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef() {
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release() {
		long cRef = InterlockedDecrement(&_cRef);
		if (cRef == 0) {
			delete this;
		}
		return cRef;
	}

	// IClassFactory
	IFACEMETHODIMP CreateInstance(IUnknown *punkOuter, REFIID riid, void **ppv) {
		return punkOuter ? CLASS_E_NOAGGREGATION : _pfnCreate(riid, ppv);
	}

	IFACEMETHODIMP LockServer(BOOL fLock) {
		if (fLock) {
			wlep::IncDllRef();
		} else {
			wlep::DecDllRef();
		}
		return S_OK;
	}

private:
	~CClassFactory() {
		wlep::DecDllRef();
	}

	long _cRef;
	PFNCREATEINSTANCE _pfnCreate;
};

// Creates a registry key (if needed) and sets the default value of the key
HRESULT CreateRegKeyAndSetValue(const REGISTRY_ENTRY *pRegistryEntry) {
	HKEY hKey;
	HRESULT hr = HRESULT_FROM_WIN32(RegCreateKeyExW(pRegistryEntry->hkeyRoot, pRegistryEntry->pszKeyName,
													0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hKey, NULL));

	if (SUCCEEDED(hr)) {
		hr = HRESULT_FROM_WIN32(RegSetValueExW(hKey, pRegistryEntry->pszValueName, 0, REG_SZ, (LPBYTE)pRegistryEntry->pszData,
			((DWORD)wcslen(pRegistryEntry->pszData) + 1) * sizeof(WCHAR)));

		RegCloseKey(hKey);
	}

	return hr;
}

// Registers this COM server
STDAPI DllRegisterServer() {
	HRESULT hr;
	WCHAR szModuleName[MAX_PATH];

	if (!GetModuleFileNameW(g_hInstance, szModuleName, ARRAYSIZE(szModuleName))) {
		return HRESULT_FROM_WIN32(GetLastError());
	} else {
		const REGISTRY_ENTRY rgRegistryEntries[] = {
			// RootKey			KeyName														   ValueName					Data
			{HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_WINLEP_THUMBNAIL_HANDLER, NULL, SZ_WINLEP_THUMBNAIL_HANDLER},
			{HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_WINLEP_THUMBNAIL_HANDLER L"\\InProcServer32", NULL, szModuleName},
			{HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_WINLEP_THUMBNAIL_HANDLER L"\\InProcServer32", L"ThreadingModel", L"Apartment"},
			{HKEY_CURRENT_USER, L"Software\\Classes\\" SZ_FILE_EXTENSION "\\ShellEx\\" SZ_THUMBNAIL_IMAGE_HANDLER_SUBKEY, NULL, SZ_CLSID_WINLEP_THUMBNAIL_HANDLER},
		};
		hr = S_OK;

		for (int i = 0; i < ARRAYSIZE(rgRegistryEntries) && SUCCEEDED(hr); i++) {
			hr = CreateRegKeyAndSetValue(&rgRegistryEntries[i]);
		}
	}
	if (SUCCEEDED(hr)) {
		// This tells the shell to invalidate the thumbnail cache. This is important because any .wlep files
		// viewed before registering this handler would otherwise show cached blank thumbnails.
		SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
	}

	return hr;
}

// Unregisters this COM server
STDAPI DllUnregisterServer() {
	HRESULT hr = S_OK;
	
	const PCWSTR rgpszKeys[] = {
		L"Software\\Classes\\CLSID\\" SZ_CLSID_WINLEP_THUMBNAIL_HANDLER,
		L"Software\\Classes\\" SZ_FILE_EXTENSION "\\ShellEx\\" SZ_THUMBNAIL_IMAGE_HANDLER_SUBKEY
	};

	// Delete the registry entries
	for (int i = 0; i < ARRAYSIZE(rgpszKeys) && SUCCEEDED(hr); i++) {
		hr = HRESULT_FROM_WIN32(RegDeleteTreeW(HKEY_CURRENT_USER, rgpszKeys[i]));
		if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) {
			// If the registry entry has already been deleted, it's S_OK.
			hr = S_OK;
		}
	}

	return hr;
}
// Only allow the DLL to be unloaded after all outstanding references have been released
STDAPI DllCanUnloadNow() {
	return dllRefCount > 0 ? S_FALSE : S_OK;
}

STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void **ppv) {
	return CClassFactory::CreateInstance(clsid, c_rgClassObjectInit, ARRAYSIZE(c_rgClassObjectInit), riid, ppv);
}

void wlep::IncDllRef() {
	// Increments the value of the specified 32-bit variable as an atomic operation.
	InterlockedIncrement(&dllRefCount);
}

void wlep::DecDllRef() {
	// Decrements the value of the specified 32-bit variable as an atomic operation.
	InterlockedDecrement(&dllRefCount);
}