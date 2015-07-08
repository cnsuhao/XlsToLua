
// ExcelTransformDlg.h : ͷ�ļ�
#include <fstream>
#include <string.h>
#include <vector>
#include <map>
#include <iostream>
using namespace std;

#pragma once

enum E_TOOL_COPYRIGHT
{
	E_TOOL_COPYRIGHT_CLIENT = 0,
	E_TOOL_COPYRIGHT_SERVER = 1,
};
typedef std::vector<CString>		VEC_STR;
typedef std::map<int, CString>		MAP_INT_STR;
typedef std::map<CString, CString>	MAP_TABLE_KEY;

// CExcelTransformDlg �Ի���
class CExcelTransformDlg : public CDialogEx
{
// ����
public:
	CExcelTransformDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EXCELTRANSFORM_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��

private:
	bool m_bFloder;
	E_TOOL_COPYRIGHT m_eCopyRight;

	bool ProcessTransform();
	bool ExcelToLua(LPCTSTR FileName);

// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedBtnInbrowse();
	afx_msg void OnBnClickedBtnSelectfolder();
	afx_msg void OnBnClickedBtnOutbrowse();
	afx_msg void OnBnClickedBtnClear();
	afx_msg void OnBnClickedBtnTransform();
};
