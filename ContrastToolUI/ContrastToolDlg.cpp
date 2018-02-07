
// ContrastToolDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ContrastTool.h"
#include "ContrastToolDlg.h"
#include "afxdialogex.h"
#include "Python.h"
#include "io.h"
#include "stringobject.h"
#include "windows.h"
#include "string.h"
#include "stdlib.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CContrastToolDlg 对话框



CContrastToolDlg::CContrastToolDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_CONTRASTTOOL_DIALOG, pParent)
	, m_editNewData(_T(""))
	, m_editOldData(_T(""))
	, m_editResult(_T(""))
	, m_filename(_T(""))
	, m_xlsortxt(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}



void CContrastToolDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_NEWDATA_EDIT, m_editNewData);
	DDX_Text(pDX, IDC_OLDDATA_EDIT, m_editOldData);
	DDX_Text(pDX, IDC_RESULT_EDIT, m_editResult);
	DDX_Text(pDX, IDC_FILENAME_EDIT, m_filename);
	DDX_Radio(pDX, IDC_XLS_RADIO, m_xlsortxt);
}

BEGIN_MESSAGE_MAP(CContrastToolDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_NEWDATA_BUTTON, &CContrastToolDlg::OnBnClickedNewdataButton)
	ON_BN_CLICKED(IDC_OLDDATA_BUTTON, &CContrastToolDlg::OnBnClickedOlddataButton)
	ON_BN_CLICKED(IDC_RESULT_BUTTON, &CContrastToolDlg::OnBnClickedResultButton)
	ON_BN_CLICKED(IDC_PRODUCE_BUTTON, &CContrastToolDlg::OnBnClickedProduceButton)
	ON_BN_CLICKED(IDC_XLS_RADIO, &CContrastToolDlg::OnBnClickedXlsRadio)
	ON_BN_CLICKED(IDC_TXT_RADIO, &CContrastToolDlg::OnBnClickedTxtRadio)
END_MESSAGE_MAP()


// CContrastToolDlg 消息处理程序

BOOL CContrastToolDlg::OnInitDialog()
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CContrastToolDlg::OnSysCommand(UINT nID, LPARAM lParam)
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
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CContrastToolDlg::OnPaint()
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
HCURSOR CContrastToolDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CContrastToolDlg::OnBnClickedNewdataButton()
{
	// TODO: 在此添加控件通知处理程序代码
	CString gReadFilePathName;
	CFileDialog fileDlg(true, _T("xlsx"), _T("*.xlsx"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("xlsx Files (*.xlsx)|*.xlsx|xls Files (*.xls)|*.xls|All Files (*.*)|*.*||"), NULL);
	if (fileDlg.DoModal() == IDOK)
	{
		gReadFilePathName = fileDlg.GetPathName();
		GetDlgItem(IDC_NEWDATA_EDIT)->SetWindowTextW(gReadFilePathName);
		CString filename = fileDlg.GetFileName();
	}
}


void CContrastToolDlg::OnBnClickedOlddataButton()
{
	// TODO: 在此添加控件通知处理程序代码
	CString gReadFilePathName;
	CFileDialog fileDlg(true, _T("xlsx"), _T("*.xlsx"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("xlsx Files (*.xlsx)|*.xlsx|xls Files (*.xls)|*.xls|All Files (*.*)|*.*||"), NULL);
	if (fileDlg.DoModal() == IDOK)
	{
		gReadFilePathName = fileDlg.GetPathName();
		GetDlgItem(IDC_OLDDATA_EDIT)->SetWindowTextW(gReadFilePathName);
		CString filename = fileDlg.GetFileName();
	}
}


void CContrastToolDlg::OnBnClickedResultButton()
{
	// TODO: 在此添加控件通知处理程序代码
	CFileFind finder;
	CString path;
	//BOOL fileExist;
	LPITEMIDLIST rootlocation;
	SHGetSpecialFolderLocation(NULL, CSIDL_DESKTOP, &rootlocation);
	if (rootlocation == NULL)
	{
		return;
	}

	BROWSEINFO bi;
	ZeroMemory(&bi, sizeof(bi));
	bi.pidlRoot = rootlocation;
	LPITEMIDLIST targetLocation = SHBrowseForFolder(&bi);
	if (targetLocation != NULL)
	{
		TCHAR targetPath[MAX_PATH];
		SHGetPathFromIDList(targetLocation, targetPath);
		GetDlgItem(IDC_RESULT_EDIT)->SetWindowTextW(targetPath);
	}
}

char* CStrToChar(CString strSrc)
{
#ifdef UNICODE  
	DWORD dwNum = WideCharToMultiByte(CP_OEMCP, NULL, strSrc.GetBuffer(0), -1, NULL, 0, NULL, FALSE);
	char *psText;
	psText = new char[dwNum];
	if (!psText)
		delete[]psText;
	WideCharToMultiByte(CP_OEMCP, NULL, strSrc.GetBuffer(0), -1, psText, dwNum, NULL, FALSE);
	return psText;
#else  
	return (LPCTSRT)strSrc;
#endif  
}

void CContrastToolDlg::OnBnClickedProduceButton()
{
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(TRUE);
	
	
	//获取exe当前目录
	char chpath[MAX_PATH];
	char *exepath = "dist/testxls.exe";
	GetModuleFileNameA(NULL, chpath, sizeof(chpath));
	chpath[strrchr(chpath, '\\') - chpath + 1] = 0;
	for(int i = 0; i <= strlen(chpath);i++)
	{
		if (chpath[i]=='\\')
		{
			chpath[i] = '/';
		}
	}
	strcat_s(chpath, exepath);
	//MessageBox((LPCTSTR)chpath);

	//执行系统命令
	char execmd[MAX_PATH] = {0};
	strcat_s(execmd, chpath);
	//取编辑框内容,将\转换为/,并转换为char *型
	m_editNewData.Replace(L"\\", L"/");
	m_editOldData.Replace(L"\\", L"/");
	m_editResult.Replace(L"\\", L"/");

	char *newdata = CStrToChar(m_editNewData);
	char *olddata = CStrToChar(m_editOldData);
	char *result = CStrToChar(m_editResult);
	char *filename = CStrToChar(m_filename);

	strcat_s(execmd, " ");
	strcat_s(execmd, newdata);
	strcat_s(execmd, " ");
	strcat_s(execmd, olddata);
	strcat_s(execmd, " ");
	strcat_s(execmd, result);
	strcat_s(execmd, "/");
	strcat_s(execmd, filename);
	if(m_xlsortxt == 0)
		strcat_s(execmd, ".xls");
	if(m_xlsortxt == 1)
		strcat_s(execmd, ".txt");
	//strcat_s(execmd, ".txt");
	//MessageBox((LPCTSTR)execmd);
	int status = system(execmd);
	if (status == -1) {
		MessageBox(_T("生成失败"), NULL, MB_ICONERROR);
	}
	else {
		MessageBox(_T("生成成功"), NULL, MB_ICONINFORMATION);
	}
	
/*
	//C++调用python函数
	Py_Initialize();
	
	if (!Py_IsInitialized())
	{
		return ;
	}

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");
	pModule = PyImport_ImportModule("testxls");      //Test001:Python文件名
	if (!pModule)
	{
		MessageBox(_T("can't find testxls.py"), NULL, MB_ICONERROR);
		return;
	}
	PyObject* pDict = PyModule_GetDict(pModule);
	if (!pDict)
	{
		MessageBox(_T("can't find dictionary"), NULL, MB_ICONERROR);
		return;
	}
	//pFunc = PyObject_GetAttrString(pModule, "comparison");  //Add:Python文件中的函数名
	pFunc = PyDict_GetItemString(pDict, "comparison");
	if (!pFunc || !PyCallable_Check(pFunc))
	{
		MessageBox(_T("can't find function [comparison]"), NULL, MB_ICONERROR);
		return;
	}

	m_editNewData.Replace(L"\\", L"/");
	m_editOldData.Replace(L"\\", L"/");
	m_editResult.Replace(L"\\", L"/");

	char *newdata = CStrToChar(m_editNewData);
	char *olddata = CStrToChar(m_editOldData);
	char *result = CStrToChar(m_editResult);

	//strncpy(a, LPCTSTR(str), strlen);
	//MessageBox((LPSTR)(LPCTSTR)str);

	PyObject *pReturn = NULL;
	pReturn = PyObject_CallFunction(pFunc, "sss", newdata, olddata, result);      //调用函数
	if (!pReturn)
	{
		MessageBox(_T("生成失败"), NULL, MB_ICONINFORMATION);
	}else
		MessageBox(_T("生成成功"), NULL, MB_ICONINFORMATION);

	//m_editResult = a;
	//MessageBox(LPCTSTR(m_editNewData));

	UpdateData(FALSE);

	Py_Finalize();
*/
}


void CContrastToolDlg::OnBnClickedXlsRadio()
{
	// TODO: 在此添加控件通知处理程序代码
	m_xlsortxt = 0;
}


void CContrastToolDlg::OnBnClickedTxtRadio()
{
	// TODO: 在此添加控件通知处理程序代码
	m_xlsortxt = 1;
}
