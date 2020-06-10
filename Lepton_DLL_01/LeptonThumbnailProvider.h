#pragma once

#include <windows.h>
#include <thumbcache.h>
#include <objidl.h>
#include <gdiplus.h>
#include <vector>
// UUID: D0385023-D398-4B09-89A2-AA1ADE3CDFE7

// Initializing The object that implements IThumbnailProvider interface must also implement IInitializeWithStream.
// The Shell calls IInitializeWithStream::Initialize with the stream of the item
// and IInitializeWithStream is the only initialization interface used when IThumbnailProvider instances are loaded out-of-proc (for isolation purposes).
class LeptonThumbnailProvider : public IInitializeWithStream, public IThumbnailProvider {
public:
	LeptonThumbnailProvider();
	~LeptonThumbnailProvider();

	// IUnknown
	// https://docs.microsoft.com/en-us/windows/win32/api/unknwn/nn-unknwn-iunknown

		// Retrieves pointers to the supported interfaces on an object.
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv) override;

	// Increments the reference count for an interface pointer to a COM object.
	// You should call this method whenever you make a copy of an interface pointer.
	IFACEMETHODIMP_(ULONG) AddRef() override;

	// When the reference count on an object reaches zero, Release must cause the interface pointer to free itself.
	// Call this method when you no longer need to use an interface pointer
	IFACEMETHODIMP_(ULONG) Release() override;

	// IThumbnailProvider
	// https://docs.microsoft.com/en-us/windows/win32/api/thumbcache/nn-thumbcache-ithumbnailprovider

	/***
	@param pStream - Pointer to an interface that represents the stream source.
	@param gfxMode - One of the STGM constants: https://docs.microsoft.com/en-us/windows/win32/stg/stgm-constants
				   - Indicates the access mode for pStream.
	***/
	IFACEMETHODIMP Initialize(IStream *pStream, DWORD gfxMode) override;

	/***
	@param cx - The maximum thumbnail size, in pixels.
	@param phbmp - Pointer to the thumbnail image handle. The image must be a DIB section and 32 bits per pixel.
	@param pdwAlpha - Pointer to the WTS_ALPHATYPE Enum. [WTSAT_UNKNOWN, WTSAT_RGB, WTSAT_ARGB]

	@return HRESULT
	If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
	***/
	IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pwdAlpha) override;

private:
	// Bytes
	static constexpr UINT m_ulWinLepPrefixSize = 2;
	// Number of bytes which represent the thumbnail size
	static constexpr UINT m_uThumbnailSizeSize = 4;
	// Number of bytes which represent the version
	static constexpr UINT m_uVersionSize = 3;

	static constexpr BYTE m_aWinLepPrefix[m_ulWinLepPrefixSize] = {0xC6, 0xD6};

	BYTE *m_pVersion;
	// Actual thumbnail size
	UINT m_uThumbnailSize;
	// Thumbnail data
	BYTE *m_pThumbnailData;

	Gdiplus::GdiplusStartupInput m_GdiplusStartupInput;
	ULONG_PTR m_pGdiplusToken;

	long m_refCount;
	// This interface supports reading and writing data to stream objects.
	// https://docs.microsoft.com/en-us/previous-versions/ms934887(v%3Dmsdn.10)
	IStream *m_pStream;

	HRESULT ValidateHeader();
	HRESULT ValidateAndReadHeader();
};
