#pragma once

class CMyIMono;

class CMyObject : public IUnknown
{
public:
	CMyObject(IUnknown * pUnkOuter);
	~CMyObject(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
	// sink events
	void						FireGratingChanged(
									long			newGrating);
	BOOL						FireBeforeMoveChange(
									double			newPosition);
	void						FireHaveAborted();
	void						FireMoveCompleted(
									long			currentPosition);
	void						FireMoveError(
									LPCTSTR			MoveError);
//	BOOL						FireQueryAllowChangeZeroOffset();
//	BOOL						FireQueryScanningStatus();
	BOOL						FireRequestChangeGrating(
									long			newGrating);
//	BOOL						FireRequestConfigFile(
//									LPTSTR			szConfigFile,
//									UINT			nBufferSize);
//	BOOL						FireRequestIniFile(
//									LPTSTR			szIniFile,
//									UINT			nBufferSize);
	HWND						FireRequestParentWindow();
	void						FireStatusMessage(
									LPCTSTR			statusMessage,
									BOOL			fAmBusy);
	BOOL						FireRequestCurrentSignal(
									double		*	signal,
									double		*	minSignal,
									double		*	maxSignal);
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIDispatch();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr);
	protected:
		HRESULT					GetCurrentWavelength(
									VARIANT		*	pVarResult);
		HRESULT					SetCurrentWavelength(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAutoGrating(
									VARIANT		*	pVarResult);
		HRESULT					SetAutoGrating(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetCurrentGrating(
									VARIANT		*	pVarResult);
		HRESULT					SetCurrentGrating(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMonoProgID(
									VARIANT		*	pVarResult);
		HRESULT					SetMonoProgID(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAmBusy(
									VARIANT		*	pVarResult);
		HRESULT					GetNumberOfGratings(
									VARIANT		*	pVarResult);
		HRESULT					GetAmInitialized(
									VARIANT		*	pVarResult);
		HRESULT					SetAmInitialized(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetGratingParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetGratingParams(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMonoParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetMonoParams(
									DISPPARAMS	*	pDispParams);
		HRESULT					IsValidPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					Home();
		HRESULT					ConvertStepsToNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					DisplayConfigValues();
		HRESULT					DoSetup();
		HRESULT					ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ReadIni(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetGratingParams(
									VARIANT		*	pVarResult);
		HRESULT					WriteConfig(
									DISPPARAMS	*	pDispParams);
		HRESULT					WriteIni(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetGratingDispersion(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					StartScan();
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IProvideClassInfo method
		STDMETHODIMP			GetClassInfo(
									ITypeInfo	**	ppTI);  
		// IProvideClassInfo2 method
		STDMETHODIMP			GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID);       
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IConnectionPointContainer methods
		STDMETHODIMP			EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum);
		STDMETHODIMP			FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	friend CImpIDispatch;
	friend CImpIConnectionPointContainer;
	friend CImpIProvideClassInfo2;
	// data members
	CImpIDispatch			*	m_pImpIDispatch;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIProvideClassInfo2	*	m_pImpIProvideClassInfo2;
	// outer unknown for aggregation
	IUnknown				*	m_pUnkOuter;
	// object reference count
	ULONG						m_cRefs;
	// connection points array
	IConnectionPoint		*	m_paConnPts[MAX_CONN_PTS];
	CMyIMono				*	m_pMyIMono;
};
