#include "stdafx.h"
#include "DateCheck.h"
#include "MyCrypto.h"
#include "MyReggie.h"

CDateCheck::CDateCheck()
{
}


CDateCheck::~CDateCheck()
{
}

BOOL CDateCheck::MyDateCheck()
{
	MyCrypto		crypt;
	CMyReggie		reggie;
	TCHAR			szString[MAX_PATH];
	TCHAR			szLastOpDate[MAX_PATH];
	TCHAR			szExpDate[MAX_PATH];
	BOOL			fSuccess = FALSE;
	BYTE		*	pByte;
	ULONG			nBytes;

	crypt.DoInit();
	// read the last op date
	if (reggie.GetStringValue(L"val1", szString, MAX_PATH))
	{
		// convert to byte
		pByte = NULL;
		crypt.TranslateRegistryString(szString, &pByte, &nBytes);
		if (NULL != pByte)
		{
			crypt.DecriptStringFromBytes_Aes(pByte, nBytes, szLastOpDate, MAX_PATH);
			CoTaskMemFree((LPVOID)pByte);
		}
	}
	// read the expiry date
	if (reggie.GetStringValue(L"val2", szString, MAX_PATH))
	{
		// convert to byte
		pByte = NULL;
		crypt.TranslateRegistryString(szString, &pByte, &nBytes);
		if (NULL != pByte)
		{
			crypt.DecriptStringFromBytes_Aes(pByte, nBytes, szExpDate, MAX_PATH);
			CoTaskMemFree((LPVOID)pByte);
		}
	}
	fSuccess = this->MyCheckDate(szLastOpDate, MAX_PATH, szExpDate);
	if (fSuccess)
	{
		reggie.SetStringValue(L"val1", szLastOpDate);
	}
	return fSuccess;
}


BOOL CDateCheck::MyCheckDate(
	LPTSTR			szLastOpDate,
	UINT			nBufferSize,
	LPCTSTR			szExpDate)
{
	HRESULT			hr;
	DATE			lastOp;			// last operated date
	DATE			expDate;		// expiry date
	LCID			lcid = 0x1009;			// Canadian English
	SYSTEMTIME		st;
	DATE			currentDate;			// current time
	BSTR			bstr;

	hr = VarDateFromStr(szLastOpDate, lcid, 0, &lastOp);
	if (FAILED(hr)) return FALSE;
	hr = VarDateFromStr(szExpDate, lcid, 0, &expDate);
	if (FAILED(hr)) return FALSE;
	// get the current date
	GetSystemTime(&st);
	SystemTimeToVariantTime(&st, &currentDate);
	// first test
	if (lastOp > (currentDate - 2))
	{
		// current date changed too much
		return FALSE;
	}
	else
	{
		// update the last operating date
		hr = VarBstrFromDate(currentDate, lcid, 0, &bstr);
		if (SUCCEEDED(hr))
		{
			StringCchCopy(szLastOpDate, nBufferSize, bstr);
			SysFreeString(bstr);
		}
		// check the expiry date
		return currentDate <= expDate;
	}
}
