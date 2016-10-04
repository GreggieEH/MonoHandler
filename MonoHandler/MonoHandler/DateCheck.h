#pragma once
class CDateCheck
{
public:
	CDateCheck();
	~CDateCheck();
	BOOL			MyDateCheck();
protected:
	BOOL			MyCheckDate(
						LPTSTR			szLastOpDate,
						UINT			nBufferSize,
						LPCTSTR			szExpDate);
};

