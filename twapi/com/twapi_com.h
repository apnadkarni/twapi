#ifndef TWAPI_COM_H
#define TWAPI_COM_H

/*
 * VMware IVix interface fragment copied from VixCOM.h
 * Not documented for twapi but used in testing and also by the tcl-vix
 * extension.
 */
typedef interface IVixHandle IVixHandle;
typedef struct IVixHandleVtbl
{
    BEGIN_INTERFACE
        
    HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
        IVixHandle * This,
        /* [in] */ REFIID riid,
        /* [annotation][iid_is][out] */ 
        void **ppvObject);
        
    ULONG ( STDMETHODCALLTYPE *AddRef )( 
        IVixHandle * This);
        
    ULONG ( STDMETHODCALLTYPE *Release )( 
        IVixHandle * This);
        
    HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
        IVixHandle * This,
        /* [out] */ UINT *pctinfo);
        
    HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
        IVixHandle * This,
        /* [in] */ UINT iTInfo,
        /* [in] */ LCID lcid,
        /* [out] */ ITypeInfo **ppTInfo);
        
    HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
        IVixHandle * This,
        /* [in] */ REFIID riid,
        /* [size_is][in] */ LPOLESTR *rgszNames,
        /* [range][in] */ UINT cNames,
        /* [in] */ LCID lcid,
        /* [size_is][out] */ DISPID *rgDispId);
        
    /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
        IVixHandle * This,
        /* [in] */ DISPID dispIdMember,
        /* [in] */ REFIID riid,
        /* [in] */ LCID lcid,
        /* [in] */ WORD wFlags,
        /* [out][in] */ DISPPARAMS *pDispParams,
        /* [out] */ VARIANT *pVarResult,
        /* [out] */ EXCEPINFO *pExcepInfo,
        /* [out] */ UINT *puArgErr);
        
    /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetHandleType )( 
        IVixHandle * This,
        /* [retval][out] */ LONG *handleType);
        
    /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetProperties )( 
        IVixHandle * This,
        /* [in] */ VARIANT propertyIDs,
        /* [out][in] */ VARIANT *propertiesArray,
        /* [retval][out] */ ULONGLONG *error);
        
    /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetPropertyType )( 
        IVixHandle * This,
        /* [in] */ LONG propertyID,
        /* [out] */ LONG *propertyType,
        /* [retval][out] */ ULONGLONG *error);
        
    /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Equals )( 
        IVixHandle * This,
        /* [in] */ IVixHandle *handle,
        /* [retval][out] */ VARIANT_BOOL *isEqual);
        
    END_INTERFACE
} IVixHandleVtbl;

interface IVixHandle
{
    CONST_VTBL struct IVixHandleVtbl *lpVtbl;
};


int Twapi_ComServerObjCmd(TwapiInterpContext *, Tcl_Interp *, int, Tcl_Obj *CONST objv[]);
int Twapi_ClassFactoryObjCmd(TwapiInterpContext *, Tcl_Interp *, int, Tcl_Obj *CONST objv[]);


#endif
