
// ExcelTransform.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelTransformApp:
// �йش����ʵ�֣������ ExcelTransform.cpp
//

class CExcelTransformApp : public CWinApp
{
public:
	CExcelTransformApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelTransformApp theApp;