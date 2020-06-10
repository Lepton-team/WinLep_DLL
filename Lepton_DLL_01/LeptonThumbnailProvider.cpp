#include "pch.h"
#include "LeptonThumbnailProvider.h"

#include <shlwapi.h>

#include <Wincrypt.h>   // For CryptStringToBinary.
#include <wincodec.h>   // Windows Imaging Codecs
#include <msxml6.h>
#include <new>

#include <sstream>
#include <string>


#include "dllmain.h"

LeptonThumbnailProvider::LeptonThumbnailProvider() : m_refCount(1), m_pStream(nullptr) {
	wlep::IncDllRef();
	// Initialized in ValidateAndReadHeader
	m_pThumbnailData = nullptr;
	m_pVersion = nullptr;
	m_uThumbnailSize = 0;

	GdiplusStartup(&m_pGdiplusToken, &m_GdiplusStartupInput, NULL);
}

LeptonThumbnailProvider::~LeptonThumbnailProvider() {
	if (m_pStream) {
		m_pStream->Release();
	}

	if (m_pThumbnailData) {
		delete[] m_pThumbnailData;
	}

	if (m_pVersion) {
		delete[] m_pVersion;
	}

	Gdiplus::GdiplusShutdown(m_pGdiplusToken);

	wlep::DecDllRef();
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

	HRESULT hr = m_pStream->QueryInterface(&m_pStream);

	return hr;
}


IFACEMETHODIMP LeptonThumbnailProvider::GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pwdAlpha) {
	if (!phbmp || !pwdAlpha) {
		return E_INVALIDARG;
	}
	// wlep thumbnail --> m_pThumbnailData  
	HRESULT hr = ValidateAndReadHeader();

	if (FAILED(hr)) {
		return hr;
	}

	*phbmp = nullptr;
	*pwdAlpha = WTS_ALPHATYPE::WTSAT_RGB;
	
	// Convert the jpg data to a stream
	IStream *pImageStream = SHCreateMemStream(m_pThumbnailData, m_uThumbnailSize);
	
	// Convert the stream to a bitmap
	Gdiplus::Bitmap *pBitmap = new Gdiplus::Bitmap(pImageStream);

	// Set the HBITMAP
	if (pBitmap->GetHBITMAP(Gdiplus::Color::Transparent, phbmp) != Gdiplus::Status::Ok) {
		pImageStream->Release();
		return E_FAIL;
	}

	pImageStream->Release();
	return S_OK;
}

HRESULT LeptonThumbnailProvider::ValidateHeader() {
	// No need to keep this as a class member
	// Since we're using it only once for validation
	BYTE *pHeader = new BYTE[m_ulWinLepPrefixSize];

	ULONG cbRead = 0;
	HRESULT hr = m_pStream->Read(pHeader, m_ulWinLepPrefixSize, &cbRead);

	if (FAILED(hr)) {
		return hr;
	}

	for (int i = 0; i < m_ulWinLepPrefixSize; i++) {
		if (pHeader[i] != m_aWinLepPrefix[i]) {
			delete[] pHeader;
			return E_FAIL;
		}
	}

	delete[] pHeader;

	return S_OK;
}

HRESULT LeptonThumbnailProvider::ValidateAndReadHeader() {
	// The first 2 bytes are already read here.
	HRESULT hr = ValidateHeader();
	if (FAILED(hr)) {
		return hr;
	}

	ULONG cbRead = 0;
	m_pVersion = new BYTE[m_uVersionSize];

	// Read the version
	hr = m_pStream->Read(m_pVersion, m_uVersionSize, &cbRead);
	if (FAILED(hr)) {
		return hr;
	}

	BYTE *pThumbnailSizeSize = new BYTE[m_uThumbnailSizeSize];
	hr = m_pStream->Read(pThumbnailSizeSize, m_uThumbnailSizeSize, &cbRead);
	if (FAILED(hr)) {
		return hr;
	}
	// Convert the read thumbnail size bytes to hex string and then convert it to a number
	// Maybe there's a better way, but I couldn't find a better and simpler way\
	// TODO: maybe memcpy() ?
	std::string hex_str = "";
	{
		for (int i = 0; i < m_uThumbnailSizeSize; i++) {
			std::stringstream stream;
			stream << std::hex << int(pThumbnailSizeSize[i]);
			hex_str += stream.str();
		}
		// We don't need it anymore
		delete[] pThumbnailSizeSize;

		// Convert the hex string to a number
		std::istringstream converter(hex_str);
		converter >> std::hex >> m_uThumbnailSize;
	} // Just to destroy istringstream

	m_pThumbnailData = new BYTE[m_uThumbnailSize];

	// Read the thumbnail data
	hr = m_pStream->Read(m_pThumbnailData, m_uThumbnailSize, &cbRead);

	return hr;
}
