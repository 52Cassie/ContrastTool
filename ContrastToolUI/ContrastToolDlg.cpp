
// ContrastToolDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CContrastToolDlg �Ի���



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


// CContrastToolDlg ��Ϣ�������

BOOL CContrastToolDlg::OnInitDialog()
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CContrastToolDlg::OnPaint()
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
HCURSOR CContrastToolDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CContrastToolDlg::OnBnClickedNewdataButton()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);
	
	
	//��ȡexe��ǰĿ¼
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

	//ִ��ϵͳ����
	char execmd[MAX_PATH] = {0};
	strcat_s(execmd, chpath);
	//ȡ�༭������,��\ת��Ϊ/,��ת��Ϊchar *��
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
		MessageBox(_T("����ʧ��"), NULL, MB_ICONERROR);
	}
	else {
		MessageBox(_T("���ɳɹ�"), NULL, MB_ICONINFORMATION);
	}
	
/*
	//C++����python����
	Py_Initialize();
	
	if (!Py_IsInitialized())
	{
		return ;
	}

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");
	pModule = PyImport_ImportModule("testxls");      //Test001:Python�ļ���
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
	//pFunc = PyObject_GetAttrString(pModule, "comparison");  //Add:Python�ļ��еĺ�����
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
	pReturn = PyObject_CallFunction(pFunc, "sss", newdata, olddata, result);      //���ú���
	if (!pReturn)
	{
		MessageBox(_T("����ʧ��"), NULL, MB_ICONINFORMATION);
	}else
		MessageBox(_T("���ɳɹ�"), NULL, MB_ICONINFORMATION);

	//m_editResult = a;
	//MessageBox(LPCTSTR(m_editNewData));

	UpdateData(FALSE);

	Py_Finalize();
*/
}


void CContrastToolDlg::OnBnClickedXlsRadio()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	m_xlsortxt = 0;
}


void CContrastToolDlg::OnBnClickedTxtRadio()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	m_xlsortxt = 1;
}
