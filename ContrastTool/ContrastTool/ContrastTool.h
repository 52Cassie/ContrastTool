
// ContrastTool.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CContrastToolApp: 
// �йش����ʵ�֣������ ContrastTool.cpp
//

class CContrastToolApp : public CWinApp
{
public:
	CContrastToolApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CContrastToolApp theApp;