#include "pch.h"
#include "LeptonPropertyHandler.h"
#include "dllmain.h"

#include <new>
#include <memory>

#include <propkey.h>
#include <propvarutil.h>
#include <shlwapi.h>
#include <strsafe.h>

using namespace lepton;

LeptonPropertyHandler::LeptonPropertyHandler() : m_refCount(1), m_pStream(nullptr), m_pStoreCache(nullptr) {
	IncDllRef();
}

LeptonPropertyHandler::~LeptonPropertyHandler() {
	DecDllRef();
}

IFACEMETHODIMP LeptonPropertyHandler::QueryInterface(REFIID riid, void** ppv) {
	static const QITAB qit[] = { QITABENT(LeptonPropertyHandler, IPropertyStore),
								 QITABENT(LeptonPropertyHandler, IInitializeWithStream),
		//The last structure in the array must have its piid member set to NULL and its dwOffset member set to 0.
		{nullptr},
#pragma warning(push)
#pragma warning(disable: 4838)
	};
#pragma warning(pop)
	return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) LeptonPropertyHandler::AddRef() {
	return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG) LeptonPropertyHandler::Release() {
	unsigned long refCount = InterlockedDecrement(&m_refCount);
	if (refCount == 0) {
		if (m_pStream) {
			m_pStream->Release();
		}

		if (m_pStoreCache) {
			m_pStoreCache->Release();
		}

		delete this;
	}
	return refCount;
}

IFACEMETHODIMP LeptonPropertyHandler::Initialize(IStream* pStream, DWORD grfMode) {
	if (m_pStream) {
		return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
	}

	HRESULT res = pStream->QueryInterface(&m_pStream);
	if (FAILED(res)) {
		return res;
	}

	res = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&m_pStoreCache));
	if (FAILED(res)) {
		return res;
	}

	// TODO: Implement

	// -----

	res = loadProperties();
	if (FAILED(res)) {
		return res;
	}

	return S_OK;
}

IFACEMETHODIMP LeptonPropertyHandler::GetCount(DWORD* pcProps) {
	if (!pcProps) {
		return E_INVALIDARG;
	}

	*pcProps = NULL;

	if (!m_pStoreCache) {
		return E_UNEXPECTED;
	}

	return m_pStoreCache->GetCount(pcProps);
}

IFACEMETHODIMP LeptonPropertyHandler::GetAt(DWORD iProp, PROPERTYKEY* pkey) {
	if (!pkey) {
		return E_INVALIDARG;
	}

	*pkey = PKEY_Null;

	if (!m_pStoreCache) {
		return E_UNEXPECTED;
	}

	return m_pStoreCache->GetAt(iProp, pkey);
}

IFACEMETHODIMP LeptonPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar) {
	if (!pPropVar) {
		return E_INVALIDARG;
	}

	PropVariantInit(pPropVar);

	if (!m_pStoreCache) {
		return E_UNEXPECTED;
	}

	return m_pStoreCache->GetValue(key, pPropVar);
}

IFACEMETHODIMP LeptonPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar) {
	// All properties are read-only.
	// According to MSDN, STG_E_INVALIDARG should be returned,
	// but that doesn't exist, so this is close enough.
	return STG_E_INVALIDPARAMETER;
}

IFACEMETHODIMP LeptonPropertyHandler::Commit() {
	// All properties are read-only.
	// According to MSDN, STG_E_INVALIDARG should be returned,
	// but that doesn't exist, so this is close enough.
	return STG_E_INVALIDPARAMETER;
}

HRESULT LeptonPropertyHandler::loadProperties() {
	// TODO: Implement

	return E_NOTIMPL;
}