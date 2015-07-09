
// ExcelTransformDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelTransform.h"
#include "ExcelTransformDlg.h"
#include "afxdialogex.h"

#include "CApplication.h"
#include "CFont0.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CChart.h"
#include "CChartObject.h"
#include "CChartObjects.h"
#include "CInterior.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CExcelTransformDlg 对话框



CExcelTransformDlg::CExcelTransformDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcelTransformDlg::IDD, pParent),
	m_bFloder(false),
	m_eCopyRight(E_TOOL_COPYRIGHT_SERVER)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelTransformDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CExcelTransformDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_INBROWSE, &CExcelTransformDlg::OnBnClickedBtnInbrowse)
	ON_BN_CLICKED(IDC_BTN_SELECTFOLDER, &CExcelTransformDlg::OnBnClickedBtnSelectfolder)
	ON_BN_CLICKED(IDC_BTN_OUTBROWSE, &CExcelTransformDlg::OnBnClickedBtnOutbrowse)
	ON_BN_CLICKED(IDC_BTN_CLEAR, &CExcelTransformDlg::OnBnClickedBtnClear)
	ON_BN_CLICKED(IDC_BTN_TRANSFORM, &CExcelTransformDlg::OnBnClickedBtnTransform)
END_MESSAGE_MAP()


// CExcelTransformDlg 消息处理程序

BOOL CExcelTransformDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	CString strMust = "\
**********************************************************************************\n\
1.xls文件格式要求: \n\
	1) 首行为文档说明；\n\
	2) 第二,三行分别表示字段说明和字段备注；\n\
	3) 第四,五行分别表示Client和Server是否拥有该字段说明；\n\
	4) 第六行为字段名，第七行开始为数据;\n\
	5) 工作表的名称必须为英文，作为TABLE名, 并且首列（KEY）不能为空；\n\
	6) 若某列数据为空（备注不填）则为TABLE名，后一列（KEY）不能为空；\n\
	7) 字段名不能以大写字母开头后跟数字，如F1、F2，这是Excel的关键字；\n\
	8) lua下标从1开始，则key的值也应该从1开始。\n\n\
2.如果提示缺少驱动程序，请完整安装Office任意版本。\n\n\
3.此工具只能用于Excel转换lua，版本号1.0； 如需更改或有不明之处请联系 13627102328 （陈）。\n\
**********************************************************************************";
	SetDlgItemText(IDC_STC_MUST, strMust);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CExcelTransformDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelTransformDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExcelTransformDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 替换字符串中特征字符串为指定字符串
bool ReplaceStr(char* pszSrc, const char* pszMatch, const char* pszReplace)
{
	if (!pszSrc || !pszMatch || !pszReplace)
	{
		return false;
	}

	char* FindPos = strstr(pszSrc, pszMatch);
	if (!FindPos)
	{
		return false;
	}

	int  StringLen;
	char szNewString[1024];
	while (FindPos)
	{
		memset(szNewString, 0, sizeof(szNewString));
		StringLen = FindPos - pszSrc;
		strncpy_s(szNewString, pszSrc, StringLen);
		strcat_s(szNewString, pszReplace);
		strcat_s(szNewString, FindPos + strlen(pszMatch));
		strcpy_s(pszSrc, sizeof(szNewString), szNewString);

		FindPos = strstr(pszSrc, pszMatch);
	}
	return true;
}

void CExcelTransformDlg::OnBnClickedBtnInbrowse()
{
	// TODO: 在此添加控件通知处理程序代码
	CString filename;//保存路径
	CFileDialog opendlg (TRUE, _T("*"), _T("*.xls"), OFN_OVERWRITEPROMPT, _T("xls Files (*.xls)|*.xls||"), NULL);
	if (opendlg.DoModal()==IDOK)
	{
		filename = opendlg.GetPathName(); 
		SetDlgItemText(IDC_EDT_SOURCE, filename);

		char szFileName[1024];
		strcpy_s(szFileName, sizeof(szFileName), (LPSTR)(LPCTSTR)filename);

		if (ReplaceStr(szFileName, PathFindExtension(filename), ".lua"))
		{
			SetDlgItemText(IDC_EDT_OUT, szFileName);
		}
	}
}


void CExcelTransformDlg::OnBnClickedBtnSelectfolder()
{
	// TODO: 在此添加控件通知处理程序代码
	BROWSEINFO bi;
	ZeroMemory(&bi, sizeof(BROWSEINFO));
	bi.lpszTitle= _T("请选择文件夹");
	bi.ulFlags = BIF_NEWDIALOGSTYLE | BIF_EDITBOX;
	LPITEMIDLIST pidl = SHBrowseForFolder(&bi);
	//LPWSTR szFolder;
	TCHAR szFolder[_MAX_PATH];
	memset(szFolder, 0, sizeof(szFolder));
	CString strFolder = _T("");
	if (pidl != NULL)
	{
		SHGetPathFromIDList(pidl, szFolder);
		//保存文件夹路径存        
		strFolder.Format(_T("%s"), szFolder);
		//afxMessageBox(strFolder);
		m_bFloder = true;
		SetDlgItemText(IDC_EDT_SOURCE, strFolder);

		SetDlgItemText(IDC_EDT_OUT, strFolder);
	}
	else
	{
		return;    
	}
}

void CExcelTransformDlg::OnBnClickedBtnOutbrowse()
{
	// TODO: 在此添加控件通知处理程序代码
	CString filename;//保存路径
	CFileDialog opendlg (TRUE, _T("*"), _T("fetout.lua"), OFN_OVERWRITEPROMPT, _T("lua Files (*.lua)|*.lua||"), NULL);
	if (opendlg.DoModal()==IDOK)
	{
		filename = opendlg.GetPathName(); 
		SetDlgItemText(IDC_EDT_OUT, filename);
	} 
}


void CExcelTransformDlg::OnBnClickedBtnClear()
{
	// TODO: 在此添加控件通知处理程序代码
	SetDlgItemText(IDC_EDT_SOURCE, "");
	SetDlgItemText(IDC_EDT_OUT, "");
	m_bFloder = false;
}


void CExcelTransformDlg::OnBnClickedBtnTransform()
{
	// TODO: 在此添加控件通知处理程序代码
	int nResult = this->MessageBox("确认要开始转换吗?\n客户端选择“是”，服务端选择“否”，选择“取消”返回", "Xls转换Lua" , MB_ICONQUESTION | MB_YESNOCANCEL);
	if (IDCANCEL == nResult) 
	{
		return;
	}
	m_eCopyRight = (IDYES == nResult) ? E_TOOL_COPYRIGHT_CLIENT : E_TOOL_COPYRIGHT_SERVER;

	// 检测文件名和文件格式是否正确
	CString cstrSource;
	GetDlgItemText(IDC_EDT_SOURCE, cstrSource);
	if (cstrSource.GetLength() <= 0 || (m_bFloder && 0 != strcmp(PathFindExtension(cstrSource), "")) || (!m_bFloder && 0 != strcmp(PathFindExtension(cstrSource), ".xls")))
	{
		MessageBox("源文件为空 或 文件夹|文件格式不正确！", "出错啦！", MB_ICONEXCLAMATION|MB_OK);
		return;
	}

	CString cstrOut;
	GetDlgItemText(IDC_EDT_OUT, cstrOut);
	if (cstrOut.GetLength() <= 0 || (m_bFloder && 0 != strcmp(PathFindExtension(cstrOut), "")) || (!m_bFloder && 0 != strcmp(PathFindExtension(cstrOut), ".lua")))
	{
		MessageBox("输出文件为空 或 文件夹|文件格式不正确！", "出错啦！", MB_ICONEXCLAMATION|MB_OK);
		return;
	} 

	// 开始转换
	if (this->ProcessTransform())
	{
		MessageBox("转换成功！", "提示", MB_ICONINFORMATION |MB_OK);
	}
	else
	{
		MessageBox("转换失败！", "出错啦！", MB_ICONERROR|MB_OK);
	}
}

// 开出转换
bool CExcelTransformDlg::ProcessTransform()
{
	// 检测源文件或文件夹是否存在
	CString cstrSource;
	GetDlgItemText(IDC_EDT_SOURCE, cstrSource);
	if (m_bFloder)
	{
		cstrSource += "\\*.xls";
	}

	CFileFind finder; 
	if (!finder.FindFile(cstrSource))
	{
		MessageBox("源文件或文件夹不存在！", "出错啦！", MB_ICONERROR|MB_OK);
		return false;
	}

	CString cstrOut;
	GetDlgItemText(IDC_EDT_OUT, cstrOut);

	// 创建Excel各个对象
	CApplication app;  // Excel程序
	CWorkbooks books;  // 工作簿集合
	CWorkbook book;    // 工作薄
	CWorksheets sheets;// 工作表集合
	CWorksheet sheet;  // 工作表
	CRange range;	   // 使用区域
	LPDISPATCH lpDisp; 
	CString sheetname; // 工作表名称
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	// 创建IDispatch接口对象
	if(!app.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox("无法启动Excel服务器", "出错啦！", MB_ICONSTOP | MB_OK);
		return false;
	}

	books.AttachDispatch(app.get_Workbooks(), TRUE);
	
	CString strFilePath;	// 真实打开的文件路径
	CString strOutFile;		// 真实输出的文件路径
	bool bWorking = true;
	do 
	{
		bWorking = !!finder.FindNextFile();
		strFilePath = finder.GetFilePath();

		// 创建输出文件
		ofstream outputFile;
		if (m_bFloder)
		{
			strOutFile.Format("%s\\%s.lua", cstrOut, finder.GetFileTitle());
		}
		else
		{
			strOutFile = cstrOut;
		}
		outputFile.open(strOutFile, std::ios::out | std::ios::trunc);

		// 文件标题注明出处
		outputFile << "-- " << strFilePath << (E_TOOL_COPYRIGHT_CLIENT == m_eCopyRight ? " (Client)" : " (Server)") << "\n\n";

		// 设置主键（工作表名做主键）
		outputFile << finder.GetFileTitle() << "= \n{\n";

		// 开始工作
		lpDisp = books.Open(strFilePath, covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional,covOptional);  
		book.AttachDispatch(lpDisp);
		sheets.AttachDispatch(book.get_Worksheets());

		// 处理所有工作表
		long lSheetsum = sheets.get_Count();
		for(int i = 1;i <= lSheetsum; i++)
		{
			sheet.AttachDispatch(sheets.get_Item(COleVariant((long)i)));
			
			CRange usedRange;
			usedRange.AttachDispatch(sheet.get_UsedRange());

			// 获取总列数
			range.AttachDispatch(usedRange.get_Columns());
			long lColNum = range.get_Count();

			// 获取总行数
			range.AttachDispatch(usedRange.get_Rows());
			long lRowNum = range.get_Count();    
			if (lColNum <= 0 || lRowNum <= 5)
			{
				continue;
			}

			COleVariant vResult;
			long lRowStart = usedRange.get_Row();	// 起始行
			long lColStart = usedRange.get_Column();// 起始列

			VEC_STR vecHeaderExplain;				// 字段说明
			VEC_STR vecHeaderRemarks;				// 字段备注
			VEC_STR vecHeader;						// 字段名
			VEC_STR vecData;						// 数据
			VEC_STR vecOldData;						// 旧数据
			int nTotalKey = 0;						// KEY的数量
			for (int i = lRowStart; i <= lRowNum; i++)
			{
				// 获取该行单元格内容
				for (int j = lColStart; j <= lColNum; j++)
				{
					range.AttachDispatch(sheet.get_Cells());
					range.AttachDispatch(range.get_Item (COleVariant((long)i),COleVariant((long)j)).pdispVal);
					vResult = range.get_Value2();
					vResult.ChangeType(VT_BSTR);

					// 第一行是文档说明（不需要检测是否有合并单元格）
					if (lRowStart == i)
					{
						outputFile << "\t-- " << (CString)vResult.bstrVal << "\n";

						outputFile << "\t" << sheet.get_Name() << "= {\n";
						break;
					}

					// 第二行是字段说明
					if (2 == i)
					{
						vecHeaderExplain.push_back((CString)vResult.bstrVal);
						continue;
					}

					// 第三行是字段备注
					if (3 == i)
					{
						vecHeaderRemarks.push_back((CString)vResult.bstrVal);
						continue;
					}

					// 第四，五行分别表示Client，Server是否拥有此字段("Y" 或 "N")
					if (4 == i || 5 == i)
					{
						if ((4 == i && E_TOOL_COPYRIGHT_CLIENT != m_eCopyRight) || (5 == i && E_TOOL_COPYRIGHT_SERVER != m_eCopyRight))
						{
							break;
						}

						vecHeader.push_back((CString)vResult.bstrVal);

						continue;
					}

					// 第六行为字段名
					if (6 == i)
					{
						if ("N" != vecHeader[j - 1])
						{
							vecHeader[j - 1] = (CString)vResult.bstrVal;
						}
						continue;
					}

					// 第七行以后为数据
					vecData.push_back((CString)vResult.bstrVal);
				}

				if (i <= 6) continue;

				// 处理需要转换的数据
				int  nKeyNum	= 1;	// 关键字数量
				int  nKeyIndex  = 0;	// 上次关键字索引
				bool bKey       = true;	// 是否处理关键字
				bool bTable     = false;// 是否处理表
				bool bCtrl      = true;	// 是否处理解析控制符
				bool bKeyForce  = false;// 是否强制处理关键字
				for (int nIndex = 0; nIndex < vecData.size(); nIndex++)
				{
					// 若此字段名没有数据或不解析则CONTINUE
					if ("" == vecHeader[nIndex] || "N" == vecHeader[nIndex])
					{
						continue;
					}

					// 文档解析控制符
					CString strCtrl;
					for (int t = 0; t < nKeyNum * 2; t++)
					{
						strCtrl += "\t";
					}

					// 处理关键字
					if (bKey)
					{
						CString strOldData;
						if (!vecOldData.empty())
						{
							strOldData = vecOldData[nIndex];
						}

						// 解析关键字（关键字不同或需要强制处理时）
						if (vecData[nIndex] != strOldData || bKeyForce)
						{
							if ("" == strOldData)
							{
								outputFile << strCtrl << "[" << vecData[nIndex] << "]= { ";	
							}
							else
							{
								// 关键字换行解析控制符
								if (bCtrl && nTotalKey > nKeyNum)
								{
									for (int n = nTotalKey; n >= nKeyNum; n--)
									{
										CString strCtrlEx = "";
										for (int m = 0; m < n + nTotalKey - 1; m++)
										{
											strCtrlEx += "\t";
										}
										outputFile << strCtrlEx << "},\n";
									}
								}
								outputFile << strCtrl << "[" << vecData[nIndex] << "]= { ";	
							}
							bTable    = true;	// 解析关键字后可以处理表了
							bKeyForce = true;	// 解析关键字后可以再次强制解析关键字了
						}
						else
						{
							bKeyForce = false;	// 未解析关键字则不能强制解析关键字了
						}
						bKey = false;			// 处理关键字后不能再次处理了
						nKeyIndex = nIndex;		// 记录上次关键字索引
						continue;
					}

					// 没有数据表示TABLE（下个字段必为改表关键字）	
					if ("" == vecData[nIndex])
					{
						// 处理表
						if (bTable)
						{
							outputFile << "\n" << strCtrl << "\t" << vecHeader[nIndex] << "= {\n";						
							outputFile << strCtrl << "\t\t" << vecHeader[nKeyIndex] << " = " << vecData[nKeyIndex] << ",\n";
							bTable = false;		// 处理表后不能再次处理了
							bCtrl  = false;		// 处理表后控制符不能再次解析
						}
						nKeyNum++;				// 增加关键字数量
						bKey = true;			// 处理表后可以处理关键字了
						continue;
					}

					// 处理一般数据
					outputFile << vecHeader[nIndex] << " = " << vecData[nIndex] << ", ";
				}
				outputFile << "},\n";			// 行结束符
				nTotalKey = nKeyNum;			// 记录总关键字数量
				vecOldData.assign(vecData.begin(), vecData.end());	// 记录当前数据
				vecData.clear();				// 清空当前数据
			}

			// 文档末 需要处理解析控制符
			for (int n = nTotalKey; n >= 1; n--)
			{
				CString strCtrlEx = "";
				for (int m = 0; m < n + nTotalKey - 1; m++)
				{
					strCtrlEx += "\t";
				}
				outputFile << strCtrlEx << "},\n";
			}
			outputFile << "\t}\n";
		}
		outputFile << "}\n";
		outputFile.flush();
		outputFile.close();

		book.Close(covOptional, COleVariant(strFilePath), covOptional);

	} while (bWorking);

	books.Close();
	books.ReleaseDispatch();
	book.ReleaseDispatch();
	sheets.ReleaseDispatch();
	sheet.ReleaseDispatch();
	range.ReleaseDispatch();
	app.ReleaseDispatch();
	app.Quit();
	return true;
}

bool CExcelTransformDlg::ExcelToLua(LPCTSTR FileName)
{
	//// 检测合并单元格
	//VARIANT varMerge = range.get_MergeCells();
	//if (varMerge.boolVal == -1)
	//{
	//	CRange rangeMerge;
	//	rangeMerge.AttachDispatch(range.get_MergeArea());
	//	rangeMerge.AttachDispatch(rangeMerge.get_Columns());
	//	int mergecol = rangeMerge.get_Count();
	//	if(mergecol > lColNum)
	//	{
	//		break;
	//	}
	//	else
	//	{
	//		continue;
	//	}
	//}
	//else if (varMerge.boolVal == 0)
	//{
	//	CString cellcontext;
	//	COleVariant vResult = range.get_Value2();
	//	cellcontext = vResult.bstrVal;
	//	outputFile << "-- " << cellcontext << "\n";
	//}
	return true;
}