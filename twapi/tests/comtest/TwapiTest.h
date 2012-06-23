// TwapiTest.h : Declaration of the CTwapiTest

#ifndef __TWAPITEST_H_
#define __TWAPITEST_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CTwapiTest
class ATL_NO_VTABLE CTwapiTest : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTwapiTest, &CLSID_TwapiTest>,
	public ISupportErrorInfo,
	public IDispatchImpl<ITwapiTest, &IID_ITwapiTest, &LIBID_COMTESTLib>
{
public:
    CTwapiTest() : ival(0), dval(0.0), saval(NULL)
    {
	bstrval = SysAllocString(L"");
	VariantInit(&variantval);
    }
    ~CTwapiTest()
    {
	SysFreeString(bstrval);
	VariantClear(&variantval);
    }

DECLARE_REGISTRY_RESOURCEID(IDR_TWAPITEST)
DECLARE_NOT_AGGREGATABLE(CTwapiTest)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CTwapiTest)
	COM_INTERFACE_ENTRY(ITwapiTest)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ITwapiTest
public:
	STDMETHOD(GetVariantType)(/*[out, retval]*/ int *pVal);
	STDMETHOD(get_IntSAProperty)(/*[out, retval]*/ SAFEARRAY **ppVal);
	STDMETHOD(put_IntSAProperty)(/*[in]*/ SAFEARRAY *pVal);
	STDMETHOD(GetUI1SA)(/*[out, retval]*/ VARIANT *varP);
	STDMETHOD(GetIntSA)(/*[out, retval]*/ VARIANT *varP);
	STDMETHOD(ThrowException)(/*[in]*/ BSTR desc);
	STDMETHOD(get_IUnknownProperty)(/*[out, retval]*/ IUnknown* *pVal);
	STDMETHOD(put_IUnknownProperty)(/*[in]*/ IUnknown* newVal);
	STDMETHOD(get_VariantProperty)(/*[out, retval]*/ VARIANT *pVal);
	STDMETHOD(put_VariantProperty)(/*[in]*/ VARIANT newVal);
	STDMETHOD(get_ScodeProperty)(/*[out, retval]*/ SCODE *pVal);
	STDMETHOD(put_ScodeProperty)(/*[in]*/ SCODE newVal);
	STDMETHOD(get_IDispatchProperty)(/*[out, retval]*/ IDispatch* *pVal);
	STDMETHOD(put_IDispatchProperty)(/*[in]*/ IDispatch* newVal);
	STDMETHOD(get_BoolProperty)(/*[out, retval]*/ BOOL *pVal);
	STDMETHOD(put_BoolProperty)(/*[in]*/ BOOL newVal);
	STDMETHOD(get_BstrProperty)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_BstrProperty)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_DateProperty)(/*[out, retval]*/ DATE *pVal);
	STDMETHOD(put_DateProperty)(/*[in]*/ DATE newVal);
	STDMETHOD(get_CurrencyProperty)(/*[out, retval]*/ CURRENCY *pVal);
	STDMETHOD(put_CurrencyProperty)(/*[in]*/ CURRENCY newVal);
	STDMETHOD(get_DoubleProperty)(/*[out, retval]*/ double *pVal);
	STDMETHOD(put_DoubleProperty)(/*[in]*/ double newVal);
	STDMETHOD(get_FloatProperty)(/*[out, retval]*/ float *pVal);
	STDMETHOD(put_FloatProperty)(/*[in]*/ float newVal);
	STDMETHOD(get_LongProperty)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_LongProperty)(/*[in]*/ long newVal);
	STDMETHOD(get_ShortProperty)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_ShortProperty)(/*[in]*/ short newVal);
private:
    int ival;
    double dval;
    BSTR  bstrval;
    bool bval;
    VARIANT variantval;
    SAFEARRAY *saval;
};

#endif //__TWAPITEST_H_
