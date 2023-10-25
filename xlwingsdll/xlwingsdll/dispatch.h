class CDispatchWrapper : public IDispatch
{
public:
    // IUnknown
	ULONG __stdcall AddRef();
	ULONG __stdcall Release();
	HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);

    // IDispatch
	HRESULT __stdcall GetTypeInfoCount(UINT* pCountTypeInfo);
	HRESULT __stdcall GetTypeInfo(UINT iTypeInfo, LCID lcid, ITypeInfo** ppITypeInfo);
	HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
	HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

	// steals a reference
	CDispatchWrapper(IDispatch* pDispatch);
	~CDispatchWrapper();

protected:
	IDispatch* pDispatch;
	ULONG cRef;
};
