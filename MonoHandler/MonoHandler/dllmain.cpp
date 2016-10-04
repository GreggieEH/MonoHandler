// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "Server.h"
#include <initguid.h>
#include "MyGuids.h"

CServer * g_pServer = NULL;

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		g_pServer = new CServer(hModule);
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		Utils_DELETE_POINTER(g_pServer);
		break;
	}
	return TRUE;
}

CServer * GetServer()
{
	return g_pServer;
}


// obtain a named interface from an object
HRESULT	GetNamedInterface(
	IUnknown			*	punk,
	LPCTSTR					szInterface,
	IDispatch			**	ppdispNamed)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	TYPEATTR			*	pTypeAttr;			// type attributes

	*ppdispNamed = NULL;
	hr = GetNamedTypeInfo(punk, szInterface, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		// have found the desired interface - get the interface
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			hr = punk->QueryInterface(pTypeAttr->guid, (LPVOID*)ppdispNamed);
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
	}
	return hr;
}

// obtain the type info for the named interface
HRESULT GetNamedTypeInfo(
	IUnknown			*	punk,
	LPCTSTR					szInterface,
	ITypeInfo			**	ppTypeInfo)
{
	HRESULT					hr;
	IProvideClassInfo	*	pProvideClassInfo;
	ITypeInfo			*	pClassInfo = NULL;			// class info
	TYPEATTR			*	pClassAttr;			// class attributes
	UINT					i;					// index over the implemented classes
	BOOL					fDone = FALSE;				// completed flag
	HREFTYPE				hRefType;			// reference type
	ITypeInfo			*	pTypeInfo = NULL;			// type info
	ITypeLib			*	pTypeLib;
	UINT					index;				// index of type info in type lib
	BSTR					bstrName;			// interface name
	LPTSTR					szName;				// interface name

	*ppTypeInfo = NULL;
	// get the class info
	hr = punk->QueryInterface(IID_IProvideClassInfo, (LPVOID*)&pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
		pProvideClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		// loop over the implemented classes
		hr = pClassInfo->GetTypeAttr(&pClassAttr);
		if (SUCCEEDED(hr))
		{
			i = 0;
			while (!fDone && i < pClassAttr->cImplTypes)
			{
				// get the interface name
				szName = NULL;
				hr = pClassInfo->GetRefTypeOfImplType(i, &hRefType);
				if (SUCCEEDED(hr))
				{
					hr = pClassInfo->GetRefTypeInfo(hRefType, &pTypeInfo);
				}
				if (SUCCEEDED(hr))
				{
					hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &index);
					if (SUCCEEDED(hr))
					{
						bstrName = NULL;
						hr = pTypeLib->GetDocumentation(index, &bstrName, NULL, NULL, NULL);
						if (SUCCEEDED(hr))
						{
							//				Utils_UnicodeToAnsi(bstrName, &szName);
							SHStrDup(bstrName, &szName);
							SysFreeString(bstrName);
						}
						pTypeLib->Release();
					}
					// check the name
					if (NULL != szName)
					{
						if (0 == lstrcmpi(szName, szInterface))
						{
							// have found the desired interface - get the interface
							*ppTypeInfo = pTypeInfo;
							pTypeInfo->AddRef();
							fDone = TRUE;
						}
						CoTaskMemFree(LPVOID(szName));
					}
					pTypeInfo->Release();
				}
				// increment the index
				if (!fDone) i++;
			}
		}
		pClassInfo->Release();
	}
	return fDone ? S_OK : E_FAIL;
}