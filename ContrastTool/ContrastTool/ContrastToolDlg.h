
// ContrastToolDlg.h : 头文件
//

#pragma once


// CContrastToolDlg 对话框
class CContrastToolDlg : public CDialogEx
{
// 构造
public:
	CContrastToolDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CONTRASTTOOL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


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
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedNewdataButton();
	afx_msg void OnBnClickedOlddataButton();
	afx_msg void OnBnClickedResultButton();
	afx_msg void OnBnClickedProduceButton();
	CString m_editNewData;
	CString m_editOldData;
	CString m_filename;
	CString m_editResult;
};
