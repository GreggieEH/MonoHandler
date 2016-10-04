#pragma once

class CMyObject;

class CMyIMono
{
public:
	CMyIMono(CMyObject * pMyObject);
	~CMyIMono(void);
	BOOL				GetProgID(
							LPTSTR			szProgID,
							UINT			nBufferSize);
	BOOL				doInit(
							LPCTSTR			szProgID);
	double				GetCurrentWavelength();
	void				SetCurrentWavelength(
							double			CurrentWavelength);
	BOOL				GetAutoGrating();
	void				SetAutoGrating(
							BOOL			AutoGrating);
	long				GetCurrentGrating();
	void				SetCurrentGrating(
							long			CurrentGrating);
	BOOL				GetAmBusy();
	long				GetNumberOfGratings();
	BOOL				GetAmInitialized();
	void				SetAmInitialized(
							BOOL			AmInitialized);
	void				GetGratingParams(
							long			grating,
							short			ParamIndex,
							VARIANT		*	Value);
	void				SetGratingParams(
							long			grating,
							short			ParamIndex,
							VARIANT		*	Value);
	void				GetMonoParams(
							short int		param,
							VARIANT		*	Value);
	void				SetMonoParams(
							short int		param,
							VARIANT		*	Value);
	BOOL				IsValidPosition(
							double			position);
	void				Home();
	BOOL				ConvertStepsToNM(	
							BOOL			fStepsToNM,
							long			gratingID,
							long		*	steps,
							double		*	nm);
	void				DisplayConfigValues();
	void				DoSetup();
	BOOL				ReadConfig(
							LPCTSTR			ConfigFile);
	BOOL				ReadIni(
							LPCTSTR			iniFile);
	BOOL				SetGratingParams();
	void				WriteConfig(
							LPCTSTR			ConfigFile);
	void				WriteIni(
							LPCTSTR			IniFile);
	// get the dispersion for the given grating
	double				CalculateDispersion(
							long			grating);
	// start scan actions
	void				ScanStart();
protected:
	// sink events
	void				OnGratingChanged(
							long			grating);
	BOOL				OnBeforeMoveChange(
							double			newPosition);
	void				OnHaveAborted();
	void				OnMoveCompleted(
							double			newPosition);
	void				OnMoveError(
							LPCTSTR			err);
//	BOOL				OnQueryAllowChangeZeroOffset();
//	BOOL				OnQueryScanningStatus();
	BOOL				OnRequestChangeGrating(
							long			newGrating);
//	BOOL				OnRequestConfigFile(
//							LPTSTR			szConfig,
//							UINT			nBufferSize);
//	BOOL				OnRequestIniFile(
//							LPTSTR			szIni,
//							UINT			nBufferSize);
	HWND				OnRequestParentWindow();
	void				OnStatusMessage(
							LPCTSTR			msg, 
							BOOL			AmBusy);
	// form the storage file name
	BOOL				FormStorageFileName(
							LPCTSTR			szIniFile,
							LPTSTR			szStorage,
							UINT			nBufferSize);
	// create storage file
	BOOL				CreateStorage(
							LPCTSTR			szFileName,
							IStorage	**	ppStg);
	// open storage file
	BOOL				OpenStorage(
							LPCTSTR			szFileName,
							IStorage	**	ppStg);
	// open object stream
	BOOL				OpenStream(
							IStorage	*	pStg,
							CLSID			clsid,
							IStream		**	ppStm);
	// create stream
	BOOL				CreateStream(
							IStorage	*	pStg,
							LPCTSTR			szStreamName,
							ULARGE_INTEGER	cbSize,
							CLSID			clsid,
							IStream		**	ppStm);
private:
	CMyObject		*	m_pMyObject;
	IDispatch		*	m_pdisp;
	IID					m_iidSink;
	DWORD				m_dwCookie;
	TCHAR				m_szProgID[MAX_PATH];
	// storage file
	IStorage		*	m_pStg;

	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CMyIMono * pMyIMono);
		~CImpISink();
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
	private:
		CMyIMono		*	m_pMyIMono;
		ULONG				m_cRefs;
		DISPID				m_dispidGratingChanged;
		DISPID				m_dispidBeforeMoveChange;
		DISPID				m_dispidHaveAborted;
		DISPID				m_dispidMoveCompleted;
		DISPID				m_dispidMoveError;
		DISPID				m_dispidRequestChangeGrating;
		DISPID				m_dispidRequestParentWindow;
		DISPID				m_dispidStatusMessage;
	};
	friend CImpISink;
};
