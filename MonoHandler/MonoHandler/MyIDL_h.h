

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Oct 04 12:49:56 2016
 */
/* Compiler settings for MyIDL.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __MyIDL_h_h__
#define __MyIDL_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IMonoHandler_FWD_DEFINED__
#define __IMonoHandler_FWD_DEFINED__
typedef interface IMonoHandler IMonoHandler;

#endif 	/* __IMonoHandler_FWD_DEFINED__ */


#ifndef ___MonoHandler_FWD_DEFINED__
#define ___MonoHandler_FWD_DEFINED__
typedef interface _MonoHandler _MonoHandler;

#endif 	/* ___MonoHandler_FWD_DEFINED__ */


#ifndef __MonoHandler_FWD_DEFINED__
#define __MonoHandler_FWD_DEFINED__

#ifdef __cplusplus
typedef class MonoHandler MonoHandler;
#else
typedef struct MonoHandler MonoHandler;
#endif /* __cplusplus */

#endif 	/* __MonoHandler_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __MonoHandler_LIBRARY_DEFINED__
#define __MonoHandler_LIBRARY_DEFINED__

/* library MonoHandler */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_MonoHandler;

#ifndef __IMonoHandler_DISPINTERFACE_DEFINED__
#define __IMonoHandler_DISPINTERFACE_DEFINED__

/* dispinterface IMonoHandler */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IMonoHandler;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("60E95CA2-0E21-45be-8258-67A576D589BD")
    IMonoHandler : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IMonoHandlerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMonoHandler * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMonoHandler * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMonoHandler * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMonoHandler * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMonoHandler * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMonoHandler * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMonoHandler * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IMonoHandlerVtbl;

    interface IMonoHandler
    {
        CONST_VTBL struct IMonoHandlerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMonoHandler_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMonoHandler_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMonoHandler_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMonoHandler_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMonoHandler_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMonoHandler_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMonoHandler_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IMonoHandler_DISPINTERFACE_DEFINED__ */


#ifndef ___MonoHandler_DISPINTERFACE_DEFINED__
#define ___MonoHandler_DISPINTERFACE_DEFINED__

/* dispinterface _MonoHandler */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__MonoHandler;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("70754D1C-F5B2-44a3-8C70-3792D641429F")
    _MonoHandler : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _MonoHandlerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _MonoHandler * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _MonoHandler * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _MonoHandler * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _MonoHandler * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _MonoHandler * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _MonoHandler * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _MonoHandler * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _MonoHandlerVtbl;

    interface _MonoHandler
    {
        CONST_VTBL struct _MonoHandlerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _MonoHandler_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _MonoHandler_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _MonoHandler_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _MonoHandler_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _MonoHandler_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _MonoHandler_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _MonoHandler_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___MonoHandler_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_MonoHandler;

#ifdef __cplusplus

class DECLSPEC_UUID("99581CC8-7944-43e8-801B-951B1624BEE5")
MonoHandler;
#endif
#endif /* __MonoHandler_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


