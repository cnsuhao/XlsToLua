
// ExcelTransformDlg.h : 头文件
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

// CExcelTransformDlg 对话框
class CExcelTransformDlg : public CDialogEx
{
// 构造
public:
	CExcelTransformDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EXCELTRANSFORM_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

private:
	bool m_bFloder;
	E_TOOL_COPYRIGHT m_eCopyRight;

	bool ProcessTransform();
	bool ExcelToLua(LPCTSTR FileName);

// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
