// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <ole2.h>
#include <olectl.h>

#include <shlwapi.h>
#include <stdio.h>
#include <tchar.h>
#include <math.h>

#include <propvarutil.h>
#include <Shobjidl.h>
#include <ShlObj.h>

#include <strsafe.h>
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\myutils.h"
#include "resource.h"

class CServer;
CServer * GetServer();

// obtain a named interface from an object
HRESULT	GetNamedInterface(
	IUnknown			*	punk,
	LPCTSTR					szInterface,
	IDispatch			**	ppdispNamed);
// obtain the type info for the named interface
HRESULT GetNamedTypeInfo(
	IUnknown			*	punk,
	LPCTSTR					szInterface,
	ITypeInfo			**	ppTypeInfo);

// definitions
#define				MY_CLSID					CLSID_MonoHandler
#define				MY_LIBID					LIBID_MonoHandler
#define				MY_IID						IID_IMonoHandler
#define				MY_IIDSINK					IID__MonoHandler

#define				MAX_CONN_PTS				1
#define				CONN_PT_CUSTOMSINK			0

#define				FRIENDLY_NAME				TEXT("MonoHandler")
#define				PROGID						TEXT("Sciencetech.MonoHandler.1")
#define				VERSIONINDEPENDENTPROGID	TEXT("Sciencetech.MonoHandler")

// mono info parameters

#define				MONO_INFO_NUMGRATINGS		100

