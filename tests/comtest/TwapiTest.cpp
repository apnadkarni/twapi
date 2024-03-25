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

STDMETHODIMP CTwapiTest::get_BoolProperty(VARIANT_BOOL *pVal)
{
    *pVal = bval;
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_BoolProperty(VARIANT_BOOL newVal)
{
    bval = newVal;
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
    SAFEARRAYBOUND bounds[3];
    SAFEARRAY *saP;
    unsigned int i,j,k,val;
    long indices[3];
    HRESULT hr;

    bounds[0].cElements = 2;
    bounds[0].lLbound = 0;
    bounds[1].cElements = 3;
    bounds[1].lLbound = 100;
    bounds[2].cElements = 4;
    bounds[2].lLbound = 1;

    saP = SafeArrayCreate(VT_I4, sizeof(bounds)/sizeof(bounds[0]), bounds);
    if (saP == NULL)
	return S_FALSE;
    for (j = 0; j < bounds[0].cElements; ++j) {
	indices[0] = bounds[0].lLbound + j;
	for (k = 0; k < bounds[1].cElements; ++k) {
	    indices[1] = bounds[1].lLbound + k;
	    for (i = 0; i < bounds[2].cElements; ++i) {
		indices[2] = bounds[2].lLbound + i;
		val = 1000 + 100*i + 10*j + k;
		hr = SafeArrayPutElement(saP, indices, &val);
		if (hr != S_OK) {
		    return hr;
		}
	    }
	}
    }

    varP->vt = VT_ARRAY | VT_I4;
    varP->parray = saP;

    return S_OK;
}

STDMETHODIMP CTwapiTest::GetUI1SA(VARIANT *varP)
{
    SAFEARRAYBOUND bounds[1];
    SAFEARRAY *saP;
    unsigned int i;
    unsigned char *p;
    HRESULT hr;

    bounds[0].cElements = 10;
    bounds[0].lLbound = 0;
    saP = SafeArrayCreate(VT_UI1, 1, bounds);
    if (saP == NULL)
	return S_FALSE;

    hr = SafeArrayAccessData(saP, (void **) &p);
    if (FAILED(hr))
	return hr;
    for (i =0; i < 10; ++i)
	*p++ = i;
    hr = SafeArrayUnaccessData(saP);
    if (FAILED(hr))
	return hr;

    varP->vt = VT_ARRAY | VT_UI1;
    varP->parray = saP;

    return S_OK;
}

STDMETHODIMP CTwapiTest::get_IntSAProperty(SAFEARRAY **saPP)
{
    HRESULT hr;
    if (saval == NULL) {
	VARIANT var;
	VariantInit(&var);
	hr = GetIntSA(&var);
	if (FAILED(hr))
	    return hr;
	saval = var.parray;
    }

    return SafeArrayCopy(saval, saPP);
}

STDMETHODIMP CTwapiTest::put_IntSAProperty(SAFEARRAY *saP)
{

    /* IMPORTANT NOTE - the safearray IDL should be defined as SAFEARRAY(int)
	to test that twapi correctly marshals VT_INT as VT_I4. If defined as
	SAFEARRAY(long), this feature is not tested as long->VT_I4 directly. */
    SAFEARRAY *sa2P;
    HRESULT hr;
    hr = SafeArrayCopy(saP, &sa2P);
    if (FAILED(hr))
	return hr;
    saval = sa2P;
    return S_OK;
}

STDMETHODIMP CTwapiTest::GetVariantType(int *pVal)
{
    *pVal = variantval.vt;
    return S_OK;
}

STDMETHODIMP CTwapiTest::GetApplicationNames(VARIANT *varP)
{
#if 0
    int vt = V_VT(varP);
    VariantClear(varP);
    V_VT(varP) = VT_I4;
    V_I4(varP) = vt;
    return S_OK;
#endif
    // Clones interface method reported by Till Immanuel
    // To invoke successfully
    //   set var [tclcast list ""]
    //   $obj GetApplicationNames var
    //   puts $var
    if ((V_VT(varP) & VT_ARRAY) == 0)
	return Error("Passed VARIANT is not a SAFEARRAY");

    VariantClear(varP);

    varP->parray = SafeArrayCreateVector(VT_BSTR, 0, 3);
    if (varP->parray == NULL)
	return S_FALSE;
    BSTR *bstrs;
    HRESULT hr = SafeArrayAccessData(varP->parray, (void **)&bstrs);
    if (FAILED(hr))
	return hr;
    bstrs[0] = SysAllocString(L"abc");
    bstrs[1] = SysAllocString(L"def");
    bstrs[2] = SysAllocString(L"ghi");
    SafeArrayUnaccessData(varP->parray);
    varP->vt = VT_ARRAY | VT_BSTR;
    return S_OK;
}


STDMETHODIMP CTwapiTest::OpenDoc6(BSTR FileName,long Type,long Options, 
    BSTR Configuration, long* Errors, long* Warnings, IDispatch** Retval)
{
    *Errors = 1;
    *Warnings = 2;
    *Retval = NULL;
    
    return S_OK;
}

STDMETHODIMP CTwapiTest::JMPSetDataVector(VARIANT vec, int *pVT)
{
	// Test for JMP application SetDataVector method
	// The method expects passed variant to be SAFEARRAY of R8 or BSTR
	// Test script should check return from this function as to
	// whether safearray type heuristics are behaving as expected.

	*pVT = vec.vt;

	return S_OK;
}

STDMETHODIMP CTwapiTest::get_UI2Property(unsigned short *pVal)
{
    *pVal = (unsigned short) ival; // TBD - generate error if does not fit
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_UI2Property(unsigned short newVal)
{
    ival = newVal;
    return S_OK;
}

STDMETHODIMP CTwapiTest::get_UI1Property(unsigned char *pVal)
{
    *pVal = (unsigned char) ival; // TBD - generate error if does not fit
    return S_OK;
}

STDMETHODIMP CTwapiTest::put_UI1Property(unsigned char newVal)
{
    ival = newVal;
    return S_OK;
}

