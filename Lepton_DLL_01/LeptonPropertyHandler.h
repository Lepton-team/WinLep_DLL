#pragma once

#include <Windows.h>
#include <propsys.h>

#include <memory>

namespace lepton {
	class LeptonPropertyHandler : public IPropertyStore, public IInitializeWithStream {
	public:
		LeptonPropertyHandler();

		// IUnknown
		IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
		IFACEMETHODIMP_(ULONG) AddRef() override;
		IFACEMETHODIMP_(ULONG) Release() override;

		// IInitializeWithStream
		IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode) override;

		// IPropertyStore
		IFACEMETHODIMP GetCount(DWORD* pcProps) override;
		IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY* pkey) override;
		IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar) override;
		IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar) override;
		IFACEMETHODIMP Commit() override;

	protected:
		~LeptonPropertyHandler();

	private:
		typedef struct fileProps {
			unsigned int width;
			unsigned int height;
		};

		ULONG m_refCount;
		IStream* m_pStream;
		IPropertyStoreCache* m_pStoreCache;
		HRESULT loadProperties();
	};
}
