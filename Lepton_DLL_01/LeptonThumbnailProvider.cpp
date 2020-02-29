#include "pch.h"
#include "LeptonThumbnailProvider.h"

#include <shlwapi.h>

#include <Wincrypt.h>   // For CryptStringToBinary.
#include <wincodec.h>   // Windows Imaging Codecs
#include <msxml6.h>
#include <new>

#include <windows.h>

#include "dllmain.h"

LeptonThumbnailProvider::LeptonThumbnailProvider() : m_refCount(1), m_pStream(nullptr) {
	lepton::IncDllRef();
}

LeptonThumbnailProvider::~LeptonThumbnailProvider() {
	if (m_pStream) {
		m_pStream->Release();
	}

	lepton::DecDllRef();
}

IFACEMETHODIMP LeptonThumbnailProvider::QueryInterface(REFIID riid, void **ppv) {
	static const QITAB qit[] = {QITABENT(LeptonThumbnailProvider, IThumbnailProvider),
								 QITABENT(LeptonThumbnailProvider, IInitializeWithStream),
		//The last structure in the array must have its piid member set to NULL and its dwOffset member set to 0.
								 {nullptr},
#pragma warning(push)
#pragma warning(disable: 4838)
	};
#pragma warning(pop)
	return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) LeptonThumbnailProvider::AddRef() {
	return InterlockedIncrement(&(m_refCount));
}

IFACEMETHODIMP_(ULONG) LeptonThumbnailProvider::Release() {
	unsigned long refCount = InterlockedDecrement(&m_refCount);
	if (refCount == 0) {
		if (m_pStream) {
			m_pStream->Release();
		}

		delete this;
	}

	return refCount;
}

IFACEMETHODIMP LeptonThumbnailProvider::Initialize(IStream *pStream, DWORD gfxMode) {
	if (m_pStream) {
		return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
	}

	HRESULT result = m_pStream->QueryInterface(&m_pStream);

	if (FAILED(result)) {
		return result;
	}

	//if (m_pStream == nullptr) {
	//	result = m_pStream->QueryInterface(&(m_pStream));
	//}

	return result;
}

IFACEMETHODIMP LeptonThumbnailProvider::GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pwdAlpha) {
	if (!phbmp || !pwdAlpha) {
		return E_INVALIDARG;
	}

	// TODO
	HRESULT res = S_OK;  // _GetBase64EncodedImageString(cx, &pszBase64EncodedImageString); // FIXME

	*phbmp = nullptr;
	*pwdAlpha = WTS_ALPHATYPE::WTSAT_ARGB;

	// TODO
	if (SUCCEEDED(res)) {
		IStream *pImageStream = nullptr;
		res = S_OK; // _GetStreamFromString(pszBase64EncodedImageString, &pImageStream); // FIXME

		// TODO
		if (SUCCEEDED(res)) {
			res = S_OK; // WICCreate32BitsPerPixelHBITMAP(pImageStream, cx, phbmp, pdwAlpha); // FIXME
			pImageStream->Release();
		}

		return S_OK;
	}
	return S_OK;
}