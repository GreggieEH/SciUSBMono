

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Mon Jan 18 22:14:07 2038
 */
/* Compiler settings for SciUsbMono.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
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
#endif /* __RPCNDR_H_VERSION__ */


#ifndef __SciUsbMono_h_h__
#define __SciUsbMono_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ISciUsbMono_FWD_DEFINED__
#define __ISciUsbMono_FWD_DEFINED__
typedef interface ISciUsbMono ISciUsbMono;

#endif 	/* __ISciUsbMono_FWD_DEFINED__ */


#ifndef ___SciUsbMono_FWD_DEFINED__
#define ___SciUsbMono_FWD_DEFINED__
typedef interface _SciUsbMono _SciUsbMono;

#endif 	/* ___SciUsbMono_FWD_DEFINED__ */


#ifndef __SciUsbMono_FWD_DEFINED__
#define __SciUsbMono_FWD_DEFINED__

#ifdef __cplusplus
typedef class SciUsbMono SciUsbMono;
#else
typedef struct SciUsbMono SciUsbMono;
#endif /* __cplusplus */

#endif 	/* __SciUsbMono_FWD_DEFINED__ */


#ifndef __IHostObject_FWD_DEFINED__
#define __IHostObject_FWD_DEFINED__
typedef interface IHostObject IHostObject;

#endif 	/* __IHostObject_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_SciUsbMono_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_SciUsbMono_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_SciUsbMono_0000_0000_v0_0_s_ifspec;


#ifndef __SciUsbMono_LIBRARY_DEFINED__
#define __SciUsbMono_LIBRARY_DEFINED__

/* library SciUsbMono */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_SciUsbMono;

#ifndef __ISciUsbMono_DISPINTERFACE_DEFINED__
#define __ISciUsbMono_DISPINTERFACE_DEFINED__

/* dispinterface ISciUsbMono */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_ISciUsbMono;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("F003B656-547D-4c79-8E7B-2BEC922BFE88")
    ISciUsbMono : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct ISciUsbMonoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISciUsbMono * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISciUsbMono * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISciUsbMono * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISciUsbMono * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISciUsbMono * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISciUsbMono * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISciUsbMono * This,
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
    } ISciUsbMonoVtbl;

    interface ISciUsbMono
    {
        CONST_VTBL struct ISciUsbMonoVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISciUsbMono_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISciUsbMono_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISciUsbMono_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISciUsbMono_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISciUsbMono_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISciUsbMono_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISciUsbMono_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __ISciUsbMono_DISPINTERFACE_DEFINED__ */


#ifndef ___SciUsbMono_DISPINTERFACE_DEFINED__
#define ___SciUsbMono_DISPINTERFACE_DEFINED__

/* dispinterface _SciUsbMono */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__SciUsbMono;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("56859167-701F-4a3e-A29B-A8BEC2C019D7")
    _SciUsbMono : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _SciUsbMonoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _SciUsbMono * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _SciUsbMono * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _SciUsbMono * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _SciUsbMono * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _SciUsbMono * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _SciUsbMono * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _SciUsbMono * This,
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
    } _SciUsbMonoVtbl;

    interface _SciUsbMono
    {
        CONST_VTBL struct _SciUsbMonoVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _SciUsbMono_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _SciUsbMono_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _SciUsbMono_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _SciUsbMono_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _SciUsbMono_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _SciUsbMono_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _SciUsbMono_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___SciUsbMono_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_SciUsbMono;

#ifdef __cplusplus

class DECLSPEC_UUID("8ADBEF97-172D-4934-8E3C-B3D7F5A870FD")
SciUsbMono;
#endif

#ifndef __IHostObject_DISPINTERFACE_DEFINED__
#define __IHostObject_DISPINTERFACE_DEFINED__

/* dispinterface IHostObject */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IHostObject;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("1A37A108-2DFB-4603-93B0-36CBF62C3304")
    IHostObject : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IHostObjectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IHostObject * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IHostObject * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IHostObject * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IHostObject * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IHostObject * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IHostObject * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IHostObject * This,
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
    } IHostObjectVtbl;

    interface IHostObject
    {
        CONST_VTBL struct IHostObjectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IHostObject_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IHostObject_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IHostObject_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IHostObject_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IHostObject_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IHostObject_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IHostObject_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IHostObject_DISPINTERFACE_DEFINED__ */

#endif /* __SciUsbMono_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


