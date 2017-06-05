#include "StdAfx.h"
#include "MyIMono.h"
#include "MyObject.h"
#include "MyGuids.h"
#include "MyCrypto.h"
#include "DateCheck.h"

CMyIMono::CMyIMono(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_pdisp(NULL),
	m_iidSink(IID_NULL),
	m_dwCookie(0),
	// storage file
	m_pStg(NULL)
{
	this->m_szProgID[0]	= '\0';
}

CMyIMono::~CMyIMono(void)
{
	if (NULL != this->m_pdisp)
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, this->m_iidSink, FALSE,
			&(this->m_dwCookie));
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
	Utils_RELEASE_INTERFACE(this->m_pStg);
}

BOOL CMyIMono::GetProgID(
							LPTSTR			szProgID,
							UINT			nBufferSize)
{
	size_t			slen;
	StringCchLength(this->m_szProgID, MAX_PATH, &slen);
	if (slen > 0)
	{
		StringCchCopy(szProgID, nBufferSize, this->m_szProgID);
		return TRUE;
	}
	else
		return FALSE;
}

BOOL CMyIMono::doInit(
							LPCTSTR			szProgID)
{
	HRESULT				hr;
//	LPOLESTR			ProgID		= NULL;
	CLSID				clsid;
	IUnknown		*	punk		= NULL;
	CImpISink		*	pSink;
	IUnknown		*	pUnkSink;
	size_t				slen, slen2;
	LPTSTR				szRem;
	CDateCheck			dateCheck;

//	if (!dateCheck.MyDateCheck())
//	{
//		MessageBox(NULL, L"Software has expired", L"SciSpec", MB_OK);
//		return FALSE;
//	}

	szRem = StrChr(szProgID, ' ');
	if (NULL != szRem)
	{
		StringCchLength(szProgID, MAX_PATH, &slen);
		StringCchLength(szRem, MAX_PATH, &slen2);
		StringCchCopyN(this->m_szProgID, MAX_PATH, szProgID, slen - slen2);
	}
	else
	{
		StringCchCopy(this->m_szProgID, MAX_PATH, szProgID);
	}

//	StringCchCopy(this->m_szProgID, MAX_PATH, L"Sciencetech.DummyMono.1");

//	Utils_AnsiToUnicode(szProgID, &ProgID);
	hr = CLSIDFromProgID(this->m_szProgID, &clsid);
//	CoTaskMemFree((LPVOID) ProgID);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (LPVOID*) &punk);
		if (FAILED(hr))
		{
			hr = CLSIDFromProgID(L"Sciencetech.DummyMono.1", &clsid);
			if (SUCCEEDED(hr))
			{
				StringCchCopy(this->m_szProgID, MAX_PATH, L"Sciencetech.DummyMono.1");
			}

		}
	}
	if (SUCCEEDED(hr))
	{
  
		hr = ::GetNamedInterface(punk, TEXT("_clsIMono"), &(this->m_pdisp));
		punk->Release();
	}
	else
	{
		MessageBox(NULL, L"Monochromator Access Failed", L"SciSpec", MB_OK);
		return FALSE;
	}
	// connect the sink
	if (SUCCEEDED(hr))
	{
		pSink	= new CImpISink(this);
		hr = pSink->QueryInterface(this->m_iidSink, (LPVOID*)&pUnkSink);
		if (SUCCEEDED(hr))
		{
			hr = Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, this->m_iidSink,
				TRUE, &(this->m_dwCookie));
			pUnkSink->Release();
		}
		else delete pSink;
	}
	return SUCCEEDED(hr);
}

double CMyIMono::GetCurrentWavelength()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("CurrentWavelength"), &dispid);
	return Utils_GetDoubleProperty(this->m_pdisp, dispid);
}

void CMyIMono::SetCurrentWavelength(
							double			CurrentWavelength)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("CurrentWavelength"), &dispid);
	Utils_SetDoubleProperty(this->m_pdisp, dispid, CurrentWavelength);
}

BOOL CMyIMono::GetAutoGrating()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AutoGrating"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

void CMyIMono::SetAutoGrating(
							BOOL			AutoGrating)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AutoGrating"), &dispid);
	Utils_SetBoolProperty(this->m_pdisp, dispid, AutoGrating);
}

long CMyIMono::GetCurrentGrating()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("CurrentGrating"), &dispid);
	return Utils_GetLongProperty(this->m_pdisp, dispid);
}

void CMyIMono::SetCurrentGrating(
							long			CurrentGrating)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("CurrentGrating"), &dispid);
	Utils_SetLongProperty(this->m_pdisp, dispid, CurrentGrating);
}

BOOL CMyIMono::GetAmBusy()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AmBusy"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

long CMyIMono::GetNumberOfGratings()
{
//	DISPID				dispid;
//	Utils_GetMemid(this->m_pdisp, TEXT("NumberOfGratings"), &dispid);
//	return Utils_GetLongProperty(this->m_pdisp, dispid);
	VARIANT			Value;
	long			numGratings		= 0;

	this->GetMonoParams(MONO_INFO_NUMGRATINGS, &Value);
	VariantToInt32(Value, &numGratings);
	return numGratings;
}

BOOL CMyIMono::GetAmInitialized()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AmInitialized"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

void CMyIMono::SetAmInitialized(
							BOOL			AmInitialized)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AmInitialized"), &dispid);
	Utils_SetBoolProperty(this->m_pdisp, dispid, AmInitialized);

	if (AmInitialized)
	{
		MyCrypto		cryp;
		cryp.RunCrypto();
	}

}

void CMyIMono::GetGratingParams(
							long			grating,
							short			ParamIndex,
							VARIANT		*	Value)
{
//	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdisp, TEXT("GratingParams"), &dispid);
	InitVariantFromInt32(grating, &avarg[1]);
	InitVariantFromInt16(ParamIndex, &avarg[0]);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYGET, avarg, 2, Value);
}

void CMyIMono::SetGratingParams(
							long			grating,
							short			ParamIndex,
							VARIANT		*	Value)
{
	return;
/*
	DISPID				dispid;
	VARIANTARG			avarg[3];
	Utils_GetMemid(this->m_pdisp, TEXT("GratingParams"), &dispid);
	InitVariantFromInt32(grating, &avarg[2]);
	Utils_ShortToVariant(ParamIndex, &avarg[1]);
	VariantInit(&avarg[0]);
	VariantCopy(&avarg[0], Value);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYPUT, avarg, 3, NULL);
	VariantClear(&avarg[0]);
*/
}

void CMyIMono::GetMonoParams(
							short int		param,
							VARIANT		*	Value)
{
	DISPID				dispid;
	VARIANTARG			varg;
	Utils_GetMemid(this->m_pdisp, TEXT("MonoParams"), &dispid);
	InitVariantFromInt16(param, &varg);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYGET, &varg, 1, Value);
}

void CMyIMono::SetMonoParams(
							short int		param,
							VARIANT		*	Value)
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdisp, TEXT("MonoParams"), &dispid);
	InitVariantFromInt16(param, &avarg[1]);
	VariantInit(&avarg[0]);
	VariantCopy(&avarg[0], Value);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYPUT, avarg, 2, NULL);
	VariantClear(&avarg[0]);
}

BOOL CMyIMono::IsValidPosition(
							double			position)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("IsValidPosition"), &dispid);
	InitVariantFromDouble(position, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

void CMyIMono::Home()
{
//	DISPID				dispid;
//	Utils_GetMemid(this->m_pdisp, TEXT("Home"), &dispid);
//	Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, NULL);
	this->SetAmInitialized(TRUE);
}

BOOL CMyIMono::ConvertStepsToNM(	
							BOOL			fStepsToNM,
							long			gratingID,
							long		*	steps,
							double		*	nm)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			avarg[4];
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("ConvertStepsToNM"), &dispid);
	InitVariantFromBoolean(fStepsToNM, &avarg[3]);
	InitVariantFromInt32(gratingID, &avarg[2]);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_I4;
	avarg[1].plVal		= steps;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_R8;
	avarg[0].pdblVal	= nm;
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 4, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

void CMyIMono::DisplayConfigValues()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("DisplayConfigValues"), &dispid);
	Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, NULL);
}

void CMyIMono::DoSetup()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("DoSetup"), &dispid);
	Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, NULL);
}

BOOL CMyIMono::ReadConfig(
							LPCTSTR			ConfigFile)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("ReadConfig"), &dispid);
	InitVariantFromString(ConfigFile, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CMyIMono::ReadIni(
							LPCTSTR			iniFile)
{
	HRESULT				hr;
	TCHAR				szStorage[MAX_PATH];
	IStorage		*	pStg;
	IPersistStreamInit*	pPersistStreamInit;
	CLSID				clsid;
	IStream			*	pStm;

	if (NULL == this->m_pdisp) return FALSE;
	this->FormStorageFileName(iniFile, szStorage, MAX_PATH);
	if (this->OpenStorage(szStorage, &pStg))
	{
		hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStreamInit);
		if (SUCCEEDED(hr))
		{
			hr = pPersistStreamInit->GetClassID(&clsid);
			if (SUCCEEDED(hr))
			{
				if (this->OpenStream(pStg, clsid, &pStm))
				{
					pPersistStreamInit->Load(pStm);
					pStm->Release();
				}
			}
			pPersistStreamInit->Release();
		}
		pStg->Release();
	}
	return TRUE;
/*
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("ReadIni"), &dispid);
	InitVariantFromString(iniFile, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
*/
}

BOOL CMyIMono::SetGratingParams()
{
	return TRUE;
/*
	HRESULT				hr;
	DISPID				dispid;
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("SetGratingParams"), &dispid);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
*/
}

void CMyIMono::WriteConfig(
							LPCTSTR			ConfigFile)
{
	DISPID				dispid;
	VARIANTARG			varg;
	Utils_GetMemid(this->m_pdisp, TEXT("WriteConfig"), &dispid);
	InitVariantFromString(ConfigFile, &varg);
	Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, NULL);
	VariantClear(&varg);
}

void CMyIMono::WriteIni(
							LPCTSTR			IniFile)
{
/*
	DISPID				dispid;
	VARIANTARG			varg;
	Utils_GetMemid(this->m_pdisp, TEXT("WriteIni"), &dispid);
	InitVariantFromString(IniFile, &varg);
	Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, NULL);
	VariantClear(&varg);
*/
	HRESULT				hr;
	TCHAR				szStorage[MAX_PATH];
	IStorage		*	pStg		= NULL;
	IPersistStreamInit*	pPersistStreamInit;
	CLSID				clsid;
	IStream			*	pStm;
	ULARGE_INTEGER		cbSize;

	if (NULL == this->m_pdisp) return;
	this->FormStorageFileName(IniFile, szStorage, MAX_PATH);
	if (!this->OpenStorage(szStorage, &pStg))
	{
		this->CreateStorage(szStorage, &pStg);
	}
	if (NULL != pStg)
	{
		hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStreamInit);
		if (SUCCEEDED(hr))
		{
			hr = pPersistStreamInit->GetClassID(&clsid);
			if (SUCCEEDED(hr))
			{
				if (this->OpenStream(pStg, clsid, &pStm))
				{
					pPersistStreamInit->Save(pStm, TRUE);
					pStm->Release();
				}
				else
				{
					pPersistStreamInit->GetSizeMax(&cbSize);
					if (this->CreateStream(pStg, TEXT("IMonoStream"), cbSize, clsid, &pStm))
					{
						pPersistStreamInit->Save(pStm, TRUE);
						pStm->Release();
					}
				}
			}
			pPersistStreamInit->Release();
		}
		pStg->Release();
	}
}

// sink events
void CMyIMono::OnGratingChanged(
							long			grating)
{
	this->m_pMyObject->FireGratingChanged(grating);
}

BOOL CMyIMono::OnBeforeMoveChange(
							double			newPosition)
{
	return this->m_pMyObject->FireBeforeMoveChange(newPosition);
}

void CMyIMono::OnHaveAborted()
{
	this->m_pMyObject->FireHaveAborted();
}

void CMyIMono::OnMoveCompleted(
							double			newPosition)
{
	this->m_pMyObject->FireMoveCompleted(newPosition);
}

void CMyIMono::OnMoveError(
							LPCTSTR			err)
{
	this->m_pMyObject->FireMoveError(err);
}

/*
BOOL CMyIMono::OnQueryAllowChangeZeroOffset()
{
	return this->m_pMyObject->FireQueryAllowChangeZeroOffset();
}

BOOL CMyIMono::OnQueryScanningStatus()
{
	return this->m_pMyObject->FireQueryScanningStatus();
}
*/

BOOL CMyIMono::OnRequestChangeGrating(
							long			newGrating)
{
	return this->m_pMyObject->FireRequestChangeGrating(newGrating);
}
/*
BOOL CMyIMono::OnRequestConfigFile(
							LPTSTR			szConfig,
							UINT			nBufferSize)
{
	return this->m_pMyObject->FireRequestConfigFile(szConfig, nBufferSize);
}

BOOL CMyIMono::OnRequestIniFile(
							LPTSTR			szIni,
							UINT			nBufferSize)
{
	return this->m_pMyObject->FireRequestIniFile(szIni, nBufferSize);
}
*/

HWND CMyIMono::OnRequestParentWindow()
{
	return this->m_pMyObject->FireRequestParentWindow();
}

void CMyIMono::OnStatusMessage(
							LPCTSTR			msg, 
							BOOL			AmBusy)
{
	this->m_pMyObject->FireStatusMessage(msg, AmBusy);
}

CMyIMono::CImpISink::CImpISink(CMyIMono * pMyIMono) :
	m_pMyIMono(pMyIMono),
	m_cRefs(0),
	m_dispidGratingChanged(DISPID_UNKNOWN),
	m_dispidBeforeMoveChange(DISPID_UNKNOWN),
	m_dispidHaveAborted(DISPID_UNKNOWN),
	m_dispidMoveCompleted(DISPID_UNKNOWN),
	m_dispidMoveError(DISPID_UNKNOWN),
	m_dispidRequestChangeGrating(DISPID_UNKNOWN),
	m_dispidRequestParentWindow(DISPID_UNKNOWN),
	m_dispidStatusMessage(DISPID_UNKNOWN)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	if (NULL == this->m_pMyIMono->m_pdisp) return;
	hr = ::GetNamedTypeInfo(this->m_pMyIMono->m_pdisp, TEXT("__clsIMono"), &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pMyIMono->m_iidSink	= pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, TEXT("GratingChanged"), &m_dispidGratingChanged);
		Utils_GetMemid(pTypeInfo, TEXT("BeforeMoveChange"), &m_dispidBeforeMoveChange);
		Utils_GetMemid(pTypeInfo, TEXT("HaveAborted"), &m_dispidHaveAborted);
		Utils_GetMemid(pTypeInfo, TEXT("MoveCompleted"), &m_dispidMoveCompleted);
		Utils_GetMemid(pTypeInfo, TEXT("MoveError"), &m_dispidMoveError);
		Utils_GetMemid(pTypeInfo, TEXT("RequestChangeGrating"), &m_dispidRequestChangeGrating);
		Utils_GetMemid(pTypeInfo, TEXT("RequestParentWindow"), &m_dispidRequestParentWindow);
		Utils_GetMemid(pTypeInfo, TEXT("StatusMessage"), &m_dispidStatusMessage);
		pTypeInfo->Release();
	}
}

CMyIMono::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CMyIMono::CImpISink::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || riid == this->m_pMyIMono->m_iidSink)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyIMono::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyIMono::CImpISink::Release()
{
	ULONG			cRefs		= --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CMyIMono::CImpISink::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CMyIMono::CImpISink::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyIMono::CImpISink::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyIMono::CImpISink::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szString		= NULL;
	TCHAR				szBuffer[MAX_PATH];
	LPOLESTR			pOleStr			= NULL;
	VariantInit(&varg);
	if (dispIdMember == m_dispidGratingChanged)
	{
		hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
		if (SUCCEEDED(hr)) 
			this->m_pMyIMono->OnGratingChanged(varg.lVal);
	}
	else if (dispIdMember == m_dispidBeforeMoveChange)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt &&
			VARIANT_TRUE == *(pDispParams->rgvarg[0].pboolVal))
		{
			hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
			if (SUCCEEDED(hr))
			{
				*(pDispParams->rgvarg[0].pboolVal) = this->m_pMyIMono->OnBeforeMoveChange(varg.dblVal) ? VARIANT_TRUE : VARIANT_FALSE;
			}
		}
	}
	else if (dispIdMember == m_dispidHaveAborted)
	{
		this->m_pMyIMono->OnHaveAborted();
	}
	else if (dispIdMember == m_dispidMoveCompleted)
	{
		hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
			this->m_pMyIMono->OnMoveCompleted(varg.dblVal);
		}
	}
	else if (dispIdMember == m_dispidMoveError)
	{
		hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
//			Utils_UnicodeToAnsi(varg.bstrVal, &szString);
			this->m_pMyIMono->OnMoveError(varg.bstrVal);
//			CoTaskMemFree((LPVOID) szString);
			VariantClear(&varg);
		}
	}
/*
	else if (dispIdMember == m_dispidQueryAllowChangeZeroOffset)
	{
		if (1 == pDispParams->cArgs && (VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt &&
			VARIANT_TRUE == *(pDispParams->rgvarg[0].pboolVal))
		{
			*(pDispParams->rgvarg[0].pboolVal) = this->m_pMyIMono->OnQueryAllowChangeZeroOffset() ? VARIANT_TRUE : VARIANT_FALSE;
		}
	}
	else if (dispIdMember == m_dispidQueryScanningStatus)
	{
		if (1 == pDispParams->cArgs && (VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt)
		{
			*(pDispParams->rgvarg[0].pboolVal) = this->m_pMyIMono->OnQueryScanningStatus() ? VARIANT_TRUE : VARIANT_FALSE;
		}
	}
*/
	else if (dispIdMember == m_dispidRequestChangeGrating)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt &&
			VARIANT_TRUE == *(pDispParams->rgvarg[0].pboolVal))
		{
			hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
			if (SUCCEEDED(hr))
			{
				*(pDispParams->rgvarg[0].pboolVal) = this->m_pMyIMono->OnRequestChangeGrating(varg.lVal) ? VARIANT_TRUE : VARIANT_FALSE;
			}
		}
	}
/*
	else if (dispIdMember == m_dispidRequestConfigFile)
	{
		if (1 == pDispParams->cArgs && (VT_BYREF | VT_BSTR) == pDispParams->rgvarg[0].vt &&
			NULL == *(pDispParams->rgvarg[0].pbstrVal))
		{
			if (this->m_pMyIMono->OnRequestConfigFile(szBuffer, MAX_PATH))
			{
				Utils_AnsiToUnicode(szBuffer, &pOleStr);
				*(pDispParams->rgvarg[0].pbstrVal) = SysAllocString(pOleStr);
				CoTaskMemFree((LPVOID) pOleStr);
			}
		}
	}
	else if (dispIdMember == m_dispidRequestIniFile)
	{
		if (1 == pDispParams->cArgs && (VT_BYREF | VT_BSTR) == pDispParams->rgvarg[0].vt &&
			NULL == *(pDispParams->rgvarg[0].pbstrVal))
		{
			if (this->m_pMyIMono->OnRequestIniFile(szBuffer, MAX_PATH))
			{
				Utils_AnsiToUnicode(szBuffer, &pOleStr);
				*(pDispParams->rgvarg[0].pbstrVal) = SysAllocString(pOleStr);
				CoTaskMemFree((LPVOID) pOleStr);
			}
		}
	}
*/
	else if (dispIdMember == m_dispidRequestParentWindow)
	{
		if (1 == pDispParams->cArgs && (VT_BYREF | VT_I4) == pDispParams->rgvarg[0].vt &&
			0 == *(pDispParams->rgvarg[0].plVal))
		{
			*(pDispParams->rgvarg[0].plVal) = (long) this->m_pMyIMono->OnRequestParentWindow();
		}
	}
	else if (dispIdMember == m_dispidStatusMessage)
	{
		hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
	//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
			StringCchCopy(szBuffer, MAX_PATH, varg.bstrVal);
	//		CoTaskMemFree((LPVOID) szString);
			VariantClear(&varg);
			hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
		}
		if (SUCCEEDED(hr))
		{
			this->m_pMyIMono->OnStatusMessage(szBuffer, VARIANT_TRUE == varg.boolVal);
		}
	}
	return S_OK;
}

// create storage file
BOOL CMyIMono::CreateStorage(
							LPCTSTR			szFileName,
							IStorage	**	ppStg)
{
	HRESULT				hr;
//	LPOLESTR			FileName	= NULL;

	*ppStg	= NULL;
//	Utils_AnsiToUnicode(szFileName, &FileName);
	hr = StgCreateDocfile(szFileName,
			STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE,
			0, ppStg);
//	CoTaskMemFree((LPVOID) FileName);
	if (SUCCEEDED(hr)) hr = WriteClassStg(*ppStg, MY_CLSID);
	return SUCCEEDED(hr);
}

// open storage file
BOOL CMyIMono::OpenStorage(
							LPCTSTR			szFileName,
							IStorage	**	ppStg)
{
	HRESULT				hr;
//	LPOLESTR			FileName	= NULL;
	CLSID				clsid;
	BOOL				fSuccess	= FALSE;

	*ppStg	= NULL;
//	Utils_AnsiToUnicode(szFileName, &FileName);
	hr = StgOpenStorage(szFileName, NULL,
		STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
		NULL, 0, ppStg);
//	CoTaskMemFree((LPVOID) FileName);
	if (SUCCEEDED(hr))
	{
		hr = ReadClassStg(*ppStg, &clsid);
	}
	if (SUCCEEDED(hr))
	{
		if (MY_CLSID == clsid)
		{
			fSuccess = TRUE;
		}
		else
		{
			Utils_RELEASE_INTERFACE(*ppStg);
		}
	}
	return fSuccess;
}

// open object stream
BOOL CMyIMono::OpenStream(
							IStorage	*	pStg,
							CLSID			clsid,
							IStream		**	ppStm)
{
	HRESULT				hr;
	IEnumSTATSTG	*	pEnum;
	STATSTG				statStg;
	IStream			*	pStm;
	CLSID				cls;
	BOOL				fDone	= FALSE;

	*ppStm		= NULL;
	hr = pStg->EnumElements(0, NULL, 0, &pEnum);
	if (SUCCEEDED(hr))
	{
		while (!fDone && S_OK == pEnum->Next(1, &statStg, NULL))
		{
			if (STGTY_STREAM == statStg.type)
			{
				hr = pStg->OpenStream(statStg.pwcsName, NULL, 
					STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, &pStm);
				if (SUCCEEDED(hr))
				{
					hr = ReadClassStm(pStm, &cls);
					if (SUCCEEDED(hr) && cls == clsid)
					{
						// have found the desired stream
						fDone = TRUE;
						*ppStm	= pStm;
						(*ppStm)->AddRef();
					}
					pStm->Release();
				}
			}
			CoTaskMemFree((LPVOID) statStg.pwcsName);
		}
	}
	return fDone;
}

// create stream
BOOL CMyIMono::CreateStream(
							IStorage	*	pStg,
							LPCTSTR			szStreamName,
							ULARGE_INTEGER	cbSize,
							CLSID			clsid,
							IStream		**	ppStm)
{
	HRESULT				hr;
//	LPOLESTR			StreamName		= NULL;
	IStream			*	pStm;
//	Utils_AnsiToUnicode(szStreamName, &StreamName);
	hr = pStg->CreateStream(szStreamName, 
		STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE, 
		0, 0, &pStm);
//	CoTaskMemFree((LPVOID) StreamName);
	if (SUCCEEDED(hr))
	{
		hr = pStm->SetSize(cbSize);
		if (SUCCEEDED(hr)) hr = WriteClassStm(pStm, clsid);
		if (SUCCEEDED(hr))
		{
			*ppStm	= pStm;
			(*ppStm)->AddRef();
		}
		pStm->Release();
	}
	return SUCCEEDED(hr);
}

// form the storage file name
BOOL CMyIMono::FormStorageFileName(
							LPCTSTR			szIniFile,
							LPTSTR			szStorage,
							UINT			nBufferSize)
{
	StringCchCopy(szStorage, nBufferSize, szIniFile);
	PathRemoveFileSpec(szStorage);
	PathAppend(szStorage, TEXT("IMono.stg"));
	return TRUE;
}

// get the dispersion for the given grating
double CMyIMono::CalculateDispersion(
	long			grating)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	double				dispersion = 1.0;
	Utils_GetMemid(this->m_pdisp, L"GetGratingDispersion", &dispid);
	InitVariantFromInt32(grating, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &dispersion);
	return dispersion;
}

// start scan actions
void CMyIMono::ScanStart()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"StartScan", &dispid);
	Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, NULL);
}