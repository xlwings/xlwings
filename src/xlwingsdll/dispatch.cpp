#include "xlwingsdll.h"

CDispatchWrapper::CDispatchWrapper(IDispatch* pDispatch)
{
	this->cRef = 1;
	this->pDispatch = pDispatch;
}

CDispatchWrapper::~CDispatchWrapper()
{
	this->pDispatch->Release();
}

HRESULT __stdcall CDispatchWrapper::QueryInterface(REFIID riid, void** ppv)
{
	if(riid == IID_IUnknown)
		*ppv = (IUnknown*) this;
	else if(riid == IID_IDispatch)
		*ppv = (IDispatch*) this;
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}

	AddRef();
	return S_OK;
}

ULONG __stdcall CDispatchWrapper::AddRef()
{
    InterlockedIncrement(&cRef);
    return cRef;
}

ULONG __stdcall CDispatchWrapper::Release()
{
    ULONG ulRefCount = InterlockedDecrement(&cRef);
    if (0 == ulRefCount)
        delete this;
    return ulRefCount;
}

HRESULT __stdcall CDispatchWrapper::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	return pDispatch->GetTypeInfoCount(pCountTypeInfo);
}

HRESULT __stdcall CDispatchWrapper::GetTypeInfo(UINT iTypeInfo, LCID lcid, ITypeInfo** ppITypeInfo)
{
	return pDispatch->GetTypeInfo(iTypeInfo, lcid, ppITypeInfo);
}

HRESULT __stdcall CDispatchWrapper::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	return pDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
}

HRESULT __stdcall CDispatchWrapper::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	// we test if pVarResult is NULL because this is what VBA sets it to if the function is called as a statement (i.e. without
	// parentheses), but this then causes the arguments to be overwritten for some reason, so we pass it a dummy result
	// variable which we then dispose of
	VARIANT result;
	VariantInit(&result);
	HRESULT hRet = pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult == NULL ? &result : pVarResult, pExcepInfo, puArgErr);
	VariantClear(&result);

	if(FAILED(hRet) && pExcepInfo->bstrDescription != NULL)
	{
		BSTR bstrOld = pExcepInfo->bstrDescription;
		std::string old;
		ToStdString(bstrOld, old);

		if(old.substr(0, 24) == "Unexpected Python Error:")
		{
			std::vector<std::string> parts;
			strsplit(old, "\n", parts, false);
			if(parts[0] == "Unexpected Python Error: Traceback (most recent call last):")
			{
				std::string neu;
				for(int k = (int) parts.size() - 1; k > 0; k--)
				{
					if(!parts[k].empty())
					{
						if(!neu.empty())
							neu += "\n";
						neu += parts[k];
					}
				}

				ToBStr(neu, pExcepInfo->bstrDescription);
				SysFreeString(bstrOld);
			}
		}
	}

	return hRet;
}
