// xls2sql.h : PROJECT_NAME ���ε{�����D�n���Y��
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�� PCH �]�t���ɮ׫e���]�t 'stdafx.h'"
#endif

#include "resource.h"		// �D�n�Ÿ�


// Cxls2sqlApp:
// �аѾ\��@�����O�� xls2sql.cpp
//

class Cxls2sqlApp : public CWinApp
{
public:
	Cxls2sqlApp();

// �мg
	public:
	virtual BOOL InitInstance();

// �{���X��@

	DECLARE_MESSAGE_MAP()
};

extern Cxls2sqlApp theApp;