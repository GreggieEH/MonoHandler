#include "StdAfx.h"
#include "MyObject.h"
#include "Server.h"
#include "dispids.h"
#include "MyGuids.h"
#include "MyIMono.h"

CMyObject::CMyObject(IUnknown * pUnkOuter) :
	m_pImpIDispatch(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIProvideClassInfo2(NULL),
	// outer unknown for aggregation
	m_pUnkOuter(pUnkOuter),
	// object reference count
	m_cRefs(0),
	m_pMyIMono(NULL)
{
	if (NULL == this->m_pUnkOuter) this->m_pUnkOuter = this;
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		this->m_paConnPts[i]	= NULL;
}

CMyObject::~CMyObject(void)
{
	Utils_DELETE_POINTER(this->m_pImpIDispatch);
	Utils_DELETE_POINTER(this->m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(this->m_pImpIProvideClassInfo2);
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
	Utils_DELETE_POINTER(this->m_pMyIMono);
}

// IUnknown methods
STDMETHODIMP CMyObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid || MY_IID == riid)
		*ppv = this->m_pImpIDispatch;
	else if (IID_IConnectionPointContainer == riid)
		*ppv = this->m_pImpIConnectionPointContainer;
	else if (IID_IProvideClassInfo == riid || IID_IProvideClassInfo2 == riid)
		*ppv = this->m_pImpIProvideClassInfo2;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CMyObject::Init()
{
	HRESULT						hr;

	this->m_pImpIDispatch					= new CImpIDispatch(this, this->m_pUnkOuter);
	this->m_pImpIConnectionPointContainer	= new CImpIConnectionPointContainer(this, this->m_pUnkOuter);
	this->m_pImpIProvideClassInfo2			= new CImpIProvideClassInfo2(this, this->m_pUnkOuter);
	this->m_pMyIMono						= new CMyIMono(this);
	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIProvideClassInfo2			&&
		NULL != this->m_pMyIMono)
	{
		hr = Utils_CreateConnectionPoint(this, MY_IIDSINK,
			 &(this->m_paConnPts[CONN_PT_CUSTOMSINK]));
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}

// sink events
void CMyObject::FireGratingChanged(
									long			newGrating)
{
	VARIANTARG			varg;
	InitVariantFromInt32(newGrating, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_GratingChanged, &varg, 1);
}

BOOL CMyObject::FireBeforeMoveChange(
									double			newPosition)
{
	VARIANTARG			avarg[2];
	VARIANT_BOOL		allowChange		= VARIANT_TRUE;
	InitVariantFromDouble(newPosition, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &allowChange;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_BeforeMoveChange, avarg, 2);
	return VARIANT_TRUE == allowChange;
}

void CMyObject::FireHaveAborted()
{
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_HaveAborted, NULL, 0);
}

void CMyObject::FireMoveCompleted(
									long			currentPosition)
{
	VARIANTARG			varg;
	InitVariantFromInt32(currentPosition, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_MoveCompleted, &varg, 1);
}

void CMyObject::FireMoveError(
									LPCTSTR			MoveError)
{
	VARIANTARG			varg;
	InitVariantFromString(MoveError, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_MoveError, &varg, 1);
	VariantClear(&varg);
}

/*
BOOL CMyObject::FireQueryAllowChangeZeroOffset()
{
	VARIANTARG			varg;
	VARIANT_BOOL		fAllow		= VARIANT_TRUE;
	VariantInit(&varg);
	varg.vt			= VT_BYREF | VT_BOOL;
	varg.pboolVal	= &fAllow;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_QueryAllowChangeZeroOffset, &varg, 1);
	return VARIANT_TRUE == fAllow;
}
*/

/*
BOOL CMyObject::FireQueryScanningStatus()
{
	VARIANTARG			varg;
	VARIANT_BOOL		fScanning	= VARIANT_FALSE;
	VariantInit(&varg);
	varg.vt			= VT_BYREF | VT_BOOL;
	varg.pboolVal	= &fScanning;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_QueryScanningStatus, &varg, 1);
	return VARIANT_TRUE == fScanning;
}
*/

BOOL CMyObject::FireRequestChangeGrating(
									long			newGrating)
{
	VARIANTARG			avarg[2];
	VARIANT_BOOL		fAllow		= VARIANT_TRUE;
	InitVariantFromInt32(newGrating, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fAllow;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestChangeGrating, avarg, 2);
	return VARIANT_TRUE == fAllow;
}

/*
BOOL CMyObject::FireRequestConfigFile(
									LPTSTR			szConfigFile,
									UINT			nBufferSize)
{
	VARIANTARG			varg;
	BSTR				bstr		= NULL;
	LPTSTR				szString	= NULL;
	BOOL				fSuccess	= FALSE;
	VariantInit(&varg);
	varg.vt			= VT_BYREF | VT_BSTR;
	varg.pbstrVal	= &bstr;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestConfigFile, &varg, 1);
	if (NULL != bstr)
	{
		Utils_UnicodeToAnsi(bstr, &szString);
		SysFreeString(bstr);
		StringCchCopy(szConfigFile, nBufferSize, szString);
		CoTaskMemFree((LPVOID) szString);
		fSuccess = TRUE;
	}
	return fSuccess;
}
*/

/*
BOOL CMyObject::FireRequestIniFile(
									LPTSTR			szIniFile,
									UINT			nBufferSize)
{
	VARIANTARG			varg;
	BSTR				bstr		= NULL;
	LPTSTR				szString	= NULL;
	BOOL				fSuccess	= FALSE;
	VariantInit(&varg);
	varg.vt			= VT_BYREF | VT_BSTR;
	varg.pbstrVal	= &bstr;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestIniFile, &varg, 1);
	if (NULL != bstr)
	{
		Utils_UnicodeToAnsi(bstr, &szString);
		SysFreeString(bstr);
		StringCchCopy(szIniFile, nBufferSize, szString);
		CoTaskMemFree((LPVOID) szString);
		fSuccess = TRUE;
	}
	return fSuccess;
}
*/

HWND CMyObject::FireRequestParentWindow()
{
	VARIANTARG		varg;
	long			lval		= 0;
	VariantInit(&varg);
	varg.vt		= VT_BYREF | VT_I4;
	varg.plVal	= &lval;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestParentWindow, &varg, 1);
	return (HWND) lval;
}

void CMyObject::FireStatusMessage(
									LPCTSTR			statusMessage,
									BOOL			fAmBusy)
{
	VARIANTARG		avarg[2];
	InitVariantFromString(statusMessage, &avarg[1]);
	InitVariantFromBoolean(fAmBusy, &avarg[0]);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_StatusMessage, avarg, 2);
	VariantClear(&avarg[1]);
}

BOOL CMyObject::FireRequestCurrentSignal(
	double		*	signal,
	double		*	minSignal,
	double		*	maxSignal)
{
	VARIANTARG		avarg[4];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	*signal = 0.0;
	*minSignal = 0.0;
	*maxSignal = 0.0;
	VariantInit(&avarg[3]);
	avarg[3].vt = VT_BYREF | VT_R8;
	avarg[3].pdblVal = signal;
	VariantInit(&avarg[2]);
	avarg[2].vt = VT_BYREF | VT_R8;
	avarg[2].pdblVal = minSignal;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_R8;
	avarg[1].pdblVal = maxSignal;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestCurrentSignal, avarg, 4);
	return VARIANT_TRUE == handled;
}

CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL)
{
}

CMyObject::CImpIDispatch::~CImpIDispatch()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo	= NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetServer()->GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(MY_IID, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	switch (dispIdMember)
	{
	case DISPID_CurrentWavelength:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetCurrentWavelength(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetCurrentWavelength(pDispParams);
		break;
	case DISPID_AutoGrating:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAutoGrating(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetAutoGrating(pDispParams);
		break;
	case DISPID_CurrentGrating:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetCurrentGrating(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetCurrentGrating(pDispParams);
		break;
	case DISPID_MonoProgID:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetMonoProgID(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetMonoProgID(pDispParams);
		break;
	case DISPID_AmBusy:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAmBusy(pVarResult);
		break;
	case DISPID_NumberOfGratings:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetNumberOfGratings(pVarResult);
		break;
	case DISPID_AmInitialized:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAmInitialized(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetAmInitialized(pDispParams);
		break;
	case DISPID_GratingParams:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetGratingParams(pDispParams, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetGratingParams(pDispParams);
		break;
	case DISPID_MonoParams:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetMonoParams(pDispParams, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetMonoParams(pDispParams);
		break;
	case DISPID_IsValidPosition:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->IsValidPosition(pDispParams, pVarResult);
		break;
	case DISPID_Home:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->Home();
		break;
	case DISPID_ConvertStepsToNM:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ConvertStepsToNM(pDispParams, pVarResult);
		break;
	case DISPID_DisplayConfigValues:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplayConfigValues();
		break;
	case DISPID_DoSetup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DoSetup();
		break;
	case DISPID_ReadConfig:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadConfig(pDispParams, pVarResult);
		break;
	case DISPID_ReadIni:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadIni(pDispParams, pVarResult);
		break;
	case DISPID_SetGratingParams:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetGratingParams(pVarResult);
		break;
	case DISPID_WriteConfig:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->WriteConfig(pDispParams);
		break;
	case DISPID_WriteIni:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->WriteIni(pDispParams);
		break;
	case DISPID_GetGratingDispersion:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetGratingDispersion(pDispParams, pVarResult);
		break;
	case DISPID_StartScan:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->StartScan();
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImpIDispatch::GetCurrentWavelength(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_pBackObj->m_pMyIMono->GetCurrentWavelength(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetCurrentWavelength(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyIMono->SetCurrentWavelength(varg.dblVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetAutoGrating(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyIMono->GetAutoGrating(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetAutoGrating(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyIMono->SetAutoGrating(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetCurrentGrating(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyIMono->GetCurrentGrating(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetCurrentGrating(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyIMono->SetCurrentGrating(varg.lVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetMonoProgID(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szProgID[MAX_PATH];
	if (this->m_pBackObj->m_pMyIMono->GetProgID(szProgID, MAX_PATH))
	{
		InitVariantFromString(szProgID, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetMonoProgID(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szProgID		= NULL;

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szProgID);
	//VariantClear(&varg);
	this->m_pBackObj->m_pMyIMono->doInit(varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szProgID);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetAmBusy(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyIMono->GetAmBusy(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetNumberOfGratings(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyIMono->GetNumberOfGratings(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetAmInitialized(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyIMono->GetAmInitialized(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetAmInitialized(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyIMono->SetAmInitialized(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetGratingParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	long				grating;
	short int			ParamIndex;
	VariantInit(&varg);
	if (NULL == pVarResult) return E_INVALIDARG;
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	grating = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	ParamIndex = varg.iVal;
	this->m_pBackObj->m_pMyIMono->GetGratingParams(grating, ParamIndex, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetGratingParams(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	long				grating;
	short int			ParamIndex;
	VariantInit(&varg);
	if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	grating = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	ParamIndex = varg.iVal;
	this->m_pBackObj->m_pMyIMono->SetGratingParams(grating, ParamIndex, &(pDispParams->rgvarg[0]));
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetMonoParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	short				paramIndex;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	paramIndex = varg.iVal;
	if (NULL == pVarResult) return E_INVALIDARG;
	this->m_pBackObj->m_pMyIMono->GetMonoParams(paramIndex, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetMonoParams(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	short				paramIndex;
	VariantInit(&varg);
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	paramIndex = varg.iVal;
	this->m_pBackObj->m_pMyIMono->SetMonoParams(paramIndex, &(pDispParams->rgvarg[0]));
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::IsValidPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	double				position;
	VariantInit(&varg);
	if (NULL == pVarResult) return E_INVALIDARG;
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	position = varg.dblVal;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyIMono->IsValidPosition(position), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::Home()
{
	this->m_pBackObj->m_pMyIMono->Home();
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ConvertStepsToNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				fStepsToNM;
	long				gratingID;
	BOOL				fSuccess		= FALSE;
	VariantInit(&varg);
	if (4 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_I4) != pDispParams->rgvarg[1].vt || 
		(VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt)
	{
		return DISP_E_TYPEMISMATCH;
	}
	hr = DispGetParam(pDispParams, 0, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fStepsToNM = VARIANT_TRUE == varg.boolVal;
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	gratingID = varg.lVal;
	fSuccess = this->m_pBackObj->m_pMyIMono->ConvertStepsToNM(fStepsToNM, gratingID,
		pDispParams->rgvarg[1].plVal, pDispParams->rgvarg[0].pdblVal);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::DisplayConfigValues()
{
	this->m_pBackObj->m_pMyIMono->DisplayConfigValues();
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::DoSetup()
{
	this->m_pBackObj->m_pMyIMono->DoSetup();
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szConfig	= NULL;
	BOOL				fSuccess	= FALSE;

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szConfig);
//	VariantClear(&varg);
	fSuccess = this->m_pBackObj->m_pMyIMono->ReadConfig(varg.bstrVal);
//	CoTaskMemFree((LPVOID) szConfig);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ReadIni(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szIni		= NULL;
	BOOL				fSuccess	= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szIni);
//	VariantClear(&varg);
	fSuccess = this->m_pBackObj->m_pMyIMono->ReadIni(varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szIni);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetGratingParams(
									VARIANT		*	pVarResult)
{
	BOOL			fSuccess		= this->m_pBackObj->m_pMyIMono->SetGratingParams();
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::WriteConfig(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szConfig	= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szConfig);
//	VariantClear(&varg);
	this->m_pBackObj->m_pMyIMono->WriteConfig(varg.bstrVal);
//	CoTaskMemFree((LPVOID) szConfig);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::WriteIni(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szIni		= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szIni);
//	VariantClear(&varg);
	this->m_pBackObj->m_pMyIMono->WriteIni(varg.bstrVal);
//	CoTaskMemFree((LPVOID) szIni);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetGratingDispersion(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	double				dispersion		= 12.0;

	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
/*

// get the dispersion for the given grating
double CMyMonoServer::CalculateDispersion(
								long			grating)
{
	double				Dispersion	= -1.0;
	VARIANT				value;
	double				focalLength;
	double				OutputAngle;
	double				pitch;				// grating pitch
	double				d;					// grating line separation in nm
	double				pi		= atan(1.0) * 4.0;

	if (grating < 0 || grating >= this->m_pMonoInfo->GetNumGratings() || 
		NULL == this->m_paGratingInfo)
		return -1.0;
	focalLength = this->m_pMonoInfo->GetFocalLength();
	OutputAngle	= this->m_pMonoInfo->GetOutputAngle();
	pitch		= this->m_paGratingInfo[grating]->GetPitch();
	if (focalLength > 0.0 && OutputAngle > 0.0 && pitch > 0.0)
	{
		d = 1.0e6 / pitch;
		Dispersion = d * cos(OutputAngle * pi / 180.0) / focalLength;
	}
	return Dispersion;
}

*/
	dispersion = this->m_pBackObj->m_pMyIMono->CalculateDispersion(varg.lVal);
	InitVariantFromDouble(dispersion, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::StartScan()
{
	this->m_pBackObj->m_pMyIMono->ScanStart();
	return S_OK;
}



CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIProvideClassInfo2::~CImpIProvideClassInfo2()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_punkOuter->Release();
}

// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
									ITypeInfo	**	ppTI) 
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

// IProvideClassInfo2 method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID)
{
	if (GUIDKIND_DEFAULT_SOURCE_DISP_IID == dwGuidKind)
	{
		*pGUID	= MY_IIDSINK;
		return S_OK;
	}
	else
	{
		*pGUID	= GUID_NULL;
		return E_INVALIDARG;	
	}
}

CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIConnectionPointContainer::~CImpIConnectionPointContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_punkOuter->Release();
}

// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts,
		ppEnum);
}

STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP)
{
	HRESULT					hr		= CONNECT_E_NOCONNECTION;
	IConnectionPoint	*	pConnPt	= NULL;
	*ppCP	= NULL;
	if (MY_IIDSINK == riid)
		pConnPt	= this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	if (NULL != pConnPt)
	{
		*ppCP = pConnPt;
		pConnPt->AddRef();
		hr = S_OK;
	}
	return hr;
}
