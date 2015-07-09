
// ExcelTransformDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CExcelTransformDlg �Ի���



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


// CExcelTransformDlg ��Ϣ�������

BOOL CExcelTransformDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	CString strMust = "\
**********************************************************************************\n\
1.xls�ļ���ʽҪ��: \n\
	1) ����Ϊ�ĵ�˵����\n\
	2) �ڶ�,���зֱ��ʾ�ֶ�˵�����ֶα�ע��\n\
	3) ����,���зֱ��ʾClient��Server�Ƿ�ӵ�и��ֶ�˵����\n\
	4) ������Ϊ�ֶ����������п�ʼΪ����;\n\
	5) ����������Ʊ���ΪӢ�ģ���ΪTABLE��, �������У�KEY������Ϊ�գ�\n\
	6) ��ĳ������Ϊ�գ���ע�����ΪTABLE������һ�У�KEY������Ϊ�գ�\n\
	7) �ֶ��������Դ�д��ĸ��ͷ������֣���F1��F2������Excel�Ĺؼ��֣�\n\
	8) lua�±��1��ʼ����key��ֵҲӦ�ô�1��ʼ��\n\n\
2.�����ʾȱ������������������װOffice����汾��\n\n\
3.�˹���ֻ������Excelת��lua���汾��1.0�� ������Ļ��в���֮������ϵ 13627102328 ���£���\n\
**********************************************************************************";
	SetDlgItemText(IDC_STC_MUST, strMust);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CExcelTransformDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CExcelTransformDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// �滻�ַ����������ַ���Ϊָ���ַ���
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString filename;//����·��
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	BROWSEINFO bi;
	ZeroMemory(&bi, sizeof(BROWSEINFO));
	bi.lpszTitle= _T("��ѡ���ļ���");
	bi.ulFlags = BIF_NEWDIALOGSTYLE | BIF_EDITBOX;
	LPITEMIDLIST pidl = SHBrowseForFolder(&bi);
	//LPWSTR szFolder;
	TCHAR szFolder[_MAX_PATH];
	memset(szFolder, 0, sizeof(szFolder));
	CString strFolder = _T("");
	if (pidl != NULL)
	{
		SHGetPathFromIDList(pidl, szFolder);
		//�����ļ���·����        
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString filename;//����·��
	CFileDialog opendlg (TRUE, _T("*"), _T("fetout.lua"), OFN_OVERWRITEPROMPT, _T("lua Files (*.lua)|*.lua||"), NULL);
	if (opendlg.DoModal()==IDOK)
	{
		filename = opendlg.GetPathName(); 
		SetDlgItemText(IDC_EDT_OUT, filename);
	} 
}


void CExcelTransformDlg::OnBnClickedBtnClear()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	SetDlgItemText(IDC_EDT_SOURCE, "");
	SetDlgItemText(IDC_EDT_OUT, "");
	m_bFloder = false;
}


void CExcelTransformDlg::OnBnClickedBtnTransform()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int nResult = this->MessageBox("ȷ��Ҫ��ʼת����?\n�ͻ���ѡ���ǡ��������ѡ�񡰷񡱣�ѡ��ȡ��������", "Xlsת��Lua" , MB_ICONQUESTION | MB_YESNOCANCEL);
	if (IDCANCEL == nResult) 
	{
		return;
	}
	m_eCopyRight = (IDYES == nResult) ? E_TOOL_COPYRIGHT_CLIENT : E_TOOL_COPYRIGHT_SERVER;

	// ����ļ������ļ���ʽ�Ƿ���ȷ
	CString cstrSource;
	GetDlgItemText(IDC_EDT_SOURCE, cstrSource);
	if (cstrSource.GetLength() <= 0 || (m_bFloder && 0 != strcmp(PathFindExtension(cstrSource), "")) || (!m_bFloder && 0 != strcmp(PathFindExtension(cstrSource), ".xls")))
	{
		MessageBox("Դ�ļ�Ϊ�� �� �ļ���|�ļ���ʽ����ȷ��", "��������", MB_ICONEXCLAMATION|MB_OK);
		return;
	}

	CString cstrOut;
	GetDlgItemText(IDC_EDT_OUT, cstrOut);
	if (cstrOut.GetLength() <= 0 || (m_bFloder && 0 != strcmp(PathFindExtension(cstrOut), "")) || (!m_bFloder && 0 != strcmp(PathFindExtension(cstrOut), ".lua")))
	{
		MessageBox("����ļ�Ϊ�� �� �ļ���|�ļ���ʽ����ȷ��", "��������", MB_ICONEXCLAMATION|MB_OK);
		return;
	} 

	// ��ʼת��
	if (this->ProcessTransform())
	{
		MessageBox("ת���ɹ���", "��ʾ", MB_ICONINFORMATION |MB_OK);
	}
	else
	{
		MessageBox("ת��ʧ�ܣ�", "��������", MB_ICONERROR|MB_OK);
	}
}

// ����ת��
bool CExcelTransformDlg::ProcessTransform()
{
	// ���Դ�ļ����ļ����Ƿ����
	CString cstrSource;
	GetDlgItemText(IDC_EDT_SOURCE, cstrSource);
	if (m_bFloder)
	{
		cstrSource += "\\*.xls";
	}

	CFileFind finder; 
	if (!finder.FindFile(cstrSource))
	{
		MessageBox("Դ�ļ����ļ��в����ڣ�", "��������", MB_ICONERROR|MB_OK);
		return false;
	}

	CString cstrOut;
	GetDlgItemText(IDC_EDT_OUT, cstrOut);

	// ����Excel��������
	CApplication app;  // Excel����
	CWorkbooks books;  // ����������
	CWorkbook book;    // ������
	CWorksheets sheets;// ��������
	CWorksheet sheet;  // ������
	CRange range;	   // ʹ������
	LPDISPATCH lpDisp; 
	CString sheetname; // ����������
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	// ����IDispatch�ӿڶ���
	if(!app.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox("�޷�����Excel������", "��������", MB_ICONSTOP | MB_OK);
		return false;
	}

	books.AttachDispatch(app.get_Workbooks(), TRUE);
	
	CString strFilePath;	// ��ʵ�򿪵��ļ�·��
	CString strOutFile;		// ��ʵ������ļ�·��
	bool bWorking = true;
	do 
	{
		bWorking = !!finder.FindNextFile();
		strFilePath = finder.GetFilePath();

		// ��������ļ�
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

		// �ļ�����ע������
		outputFile << "-- " << strFilePath << (E_TOOL_COPYRIGHT_CLIENT == m_eCopyRight ? " (Client)" : " (Server)") << "\n\n";

		// ��������������������������
		outputFile << finder.GetFileTitle() << "= \n{\n";

		// ��ʼ����
		lpDisp = books.Open(strFilePath, covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional,covOptional);  
		book.AttachDispatch(lpDisp);
		sheets.AttachDispatch(book.get_Worksheets());

		// �������й�����
		long lSheetsum = sheets.get_Count();
		for(int i = 1;i <= lSheetsum; i++)
		{
			sheet.AttachDispatch(sheets.get_Item(COleVariant((long)i)));
			
			CRange usedRange;
			usedRange.AttachDispatch(sheet.get_UsedRange());

			// ��ȡ������
			range.AttachDispatch(usedRange.get_Columns());
			long lColNum = range.get_Count();

			// ��ȡ������
			range.AttachDispatch(usedRange.get_Rows());
			long lRowNum = range.get_Count();    
			if (lColNum <= 0 || lRowNum <= 5)
			{
				continue;
			}

			COleVariant vResult;
			long lRowStart = usedRange.get_Row();	// ��ʼ��
			long lColStart = usedRange.get_Column();// ��ʼ��

			VEC_STR vecHeaderExplain;				// �ֶ�˵��
			VEC_STR vecHeaderRemarks;				// �ֶα�ע
			VEC_STR vecHeader;						// �ֶ���
			VEC_STR vecData;						// ����
			VEC_STR vecOldData;						// ������
			int nTotalKey = 0;						// KEY������
			for (int i = lRowStart; i <= lRowNum; i++)
			{
				// ��ȡ���е�Ԫ������
				for (int j = lColStart; j <= lColNum; j++)
				{
					range.AttachDispatch(sheet.get_Cells());
					range.AttachDispatch(range.get_Item (COleVariant((long)i),COleVariant((long)j)).pdispVal);
					vResult = range.get_Value2();
					vResult.ChangeType(VT_BSTR);

					// ��һ�����ĵ�˵��������Ҫ����Ƿ��кϲ���Ԫ��
					if (lRowStart == i)
					{
						outputFile << "\t-- " << (CString)vResult.bstrVal << "\n";

						outputFile << "\t" << sheet.get_Name() << "= {\n";
						break;
					}

					// �ڶ������ֶ�˵��
					if (2 == i)
					{
						vecHeaderExplain.push_back((CString)vResult.bstrVal);
						continue;
					}

					// ���������ֶα�ע
					if (3 == i)
					{
						vecHeaderRemarks.push_back((CString)vResult.bstrVal);
						continue;
					}

					// ���ģ����зֱ��ʾClient��Server�Ƿ�ӵ�д��ֶ�("Y" �� "N")
					if (4 == i || 5 == i)
					{
						if ((4 == i && E_TOOL_COPYRIGHT_CLIENT != m_eCopyRight) || (5 == i && E_TOOL_COPYRIGHT_SERVER != m_eCopyRight))
						{
							break;
						}

						vecHeader.push_back((CString)vResult.bstrVal);

						continue;
					}

					// ������Ϊ�ֶ���
					if (6 == i)
					{
						if ("N" != vecHeader[j - 1])
						{
							vecHeader[j - 1] = (CString)vResult.bstrVal;
						}
						continue;
					}

					// �������Ժ�Ϊ����
					vecData.push_back((CString)vResult.bstrVal);
				}

				if (i <= 6) continue;

				// ������Ҫת��������
				int  nKeyNum	= 1;	// �ؼ�������
				int  nKeyIndex  = 0;	// �ϴιؼ�������
				bool bKey       = true;	// �Ƿ���ؼ���
				bool bTable     = false;// �Ƿ����
				bool bCtrl      = true;	// �Ƿ���������Ʒ�
				bool bKeyForce  = false;// �Ƿ�ǿ�ƴ���ؼ���
				for (int nIndex = 0; nIndex < vecData.size(); nIndex++)
				{
					// �����ֶ���û�����ݻ򲻽�����CONTINUE
					if ("" == vecHeader[nIndex] || "N" == vecHeader[nIndex])
					{
						continue;
					}

					// �ĵ��������Ʒ�
					CString strCtrl;
					for (int t = 0; t < nKeyNum * 2; t++)
					{
						strCtrl += "\t";
					}

					// ����ؼ���
					if (bKey)
					{
						CString strOldData;
						if (!vecOldData.empty())
						{
							strOldData = vecOldData[nIndex];
						}

						// �����ؼ��֣��ؼ��ֲ�ͬ����Ҫǿ�ƴ���ʱ��
						if (vecData[nIndex] != strOldData || bKeyForce)
						{
							if ("" == strOldData)
							{
								outputFile << strCtrl << "[" << vecData[nIndex] << "]= { ";	
							}
							else
							{
								// �ؼ��ֻ��н������Ʒ�
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
							bTable    = true;	// �����ؼ��ֺ���Դ������
							bKeyForce = true;	// �����ؼ��ֺ�����ٴ�ǿ�ƽ����ؼ�����
						}
						else
						{
							bKeyForce = false;	// δ�����ؼ�������ǿ�ƽ����ؼ�����
						}
						bKey = false;			// ����ؼ��ֺ����ٴδ�����
						nKeyIndex = nIndex;		// ��¼�ϴιؼ�������
						continue;
					}

					// û�����ݱ�ʾTABLE���¸��ֶα�Ϊ�ı�ؼ��֣�	
					if ("" == vecData[nIndex])
					{
						// �����
						if (bTable)
						{
							outputFile << "\n" << strCtrl << "\t" << vecHeader[nIndex] << "= {\n";						
							outputFile << strCtrl << "\t\t" << vecHeader[nKeyIndex] << " = " << vecData[nKeyIndex] << ",\n";
							bTable = false;		// ���������ٴδ�����
							bCtrl  = false;		// ��������Ʒ������ٴν���
						}
						nKeyNum++;				// ���ӹؼ�������
						bKey = true;			// ��������Դ���ؼ�����
						continue;
					}

					// ����һ������
					outputFile << vecHeader[nIndex] << " = " << vecData[nIndex] << ", ";
				}
				outputFile << "},\n";			// �н�����
				nTotalKey = nKeyNum;			// ��¼�ܹؼ�������
				vecOldData.assign(vecData.begin(), vecData.end());	// ��¼��ǰ����
				vecData.clear();				// ��յ�ǰ����
			}

			// �ĵ�ĩ ��Ҫ����������Ʒ�
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
	//// ���ϲ���Ԫ��
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