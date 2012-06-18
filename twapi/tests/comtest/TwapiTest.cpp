// TwapiTest.cpp : Implementation of CTwapiTest
#include "stdafx.h"
#include "Comtest.h"
#include "TwapiTest.h"

/////////////////////////////////////////////////////////////////////////////
// CTwapiTest

STDMETHODIMP CTwapiTest::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ITwapiTest
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CTwapiTest::get_ShortProperty(short *pVal)
{
    *pVal = (short) ival; // TBD - generate error if does not fit in short
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_ShortProperty(short newVal)
{
    ival = newVal;
    return S_OK;
}

STDMETHODIMP CTwapiTest::get_LongProperty(long *pVal)
{
    *pVal = ival;
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_LongProperty(long newVal)
{
    ival = newVal;
    return S_OK;
}

STDMETHODIMP CTwapiTest::get_FloatProperty(float *pVal)
{
    *pVal = (float) dval;
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_FloatProperty(float newVal)
{
    dval = newVal;
    return S_OK;
}

STDMETHODIMP CTwapiTest::get_DoubleProperty(double *pVal)
{
    *pVal = (float) dval;
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_DoubleProperty(double newVal)
{
    dval = newVal;
    return S_OK;
}

STDMETHODIMP CTwapiTest::get_CurrencyProperty(CURRENCY *pVal)
{
	// TODO: Add your implementation code here

	return S_FALSE;
}

STDMETHODIMP CTwapiTest::put_CurrencyProperty(CURRENCY newVal)
{
	// TODO: Add your implementation code here

	return S_FALSE;
}

STDMETHODIMP CTwapiTest::get_DateProperty(DATE *pVal)
{
	// TODO: Add your implementation code here

	return S_FALSE;
}

STDMETHODIMP CTwapiTest::put_DateProperty(DATE newVal)
{
	// TODO: Add your implementation code here

	return S_FALSE;
}

STDMETHODIMP CTwapiTest::get_BstrProperty(BSTR *pVal)
{
    *pVal = SysAllocStringLen(bstrval, SysStringLen(bstrval));
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_BstrProperty(BSTR newVal)
{
    bstrval = SysAllocStringLen(newVal, SysStringLen(newVal));
    return S_OK;
}

STDMETHODIMP CTwapiTest::get_BoolProperty(BOOL *pVal)
{
    *pVal = bval ? 1 : 0;
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_BoolProperty(BOOL newVal)
{
    bval = newVal ? 1 : 0;
    return S_OK;
}

STDMETHODIMP CTwapiTest::get_IDispatchProperty(IDispatch **pVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CTwapiTest::put_IDispatchProperty(IDispatch *newVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CTwapiTest::get_ScodeProperty(SCODE *pVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CTwapiTest::put_ScodeProperty(SCODE newVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CTwapiTest::get_VariantProperty(VARIANT *pVal)
{
    VariantInit(pVal);
    return VariantCopy(pVal, &variantval);
}

STDMETHODIMP CTwapiTest::put_VariantProperty(VARIANT newVal)
{
    return VariantCopy(&variantval, &newVal);
}

STDMETHODIMP CTwapiTest::get_IUnknownProperty(IUnknown **pVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CTwapiTest::put_IUnknownProperty(IUnknown *newVal)
{
	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CTwapiTest::ThrowException(BSTR desc)
{
    return Error(desc);
}

STDMETHODIMP CTwapiTest::GetIntSA(VARIANT *varP)
{
    SAFEARRAYBOUND bounds[2];
    SAFEARRAY *saP;
    unsigned int j,k,val;
    long indices[2];
    HRESULT hr;

    bounds[0].cElements = 2;
    bounds[0].lLbound = 0;
    bounds[1].cElements = 3;
    bounds[1].lLbound = 100;

    saP = SafeArrayCreate(VT_I4, sizeof(bounds)/sizeof(bounds[0]), bounds);
    if (saP == NULL)
	return S_FALSE;

    for (j = 0; j < bounds[0].cElements; ++j) {
	indices[0] = bounds[0].lLbound + j;
	for (k = 0; k < bounds[1].cElements; ++k) {
	    indices[1] = bounds[1].lLbound + k;
	    val = 10*j + k;
	    hr = SafeArrayPutElement(saP, indices, &val);
	    if (hr != S_OK) {
	        return hr;
	    }
	}
    }

    varP->vt = VT_ARRAY | VT_I4;
    varP->parray = saP;

    return S_OK;
}
