
// ContrastToolDlg.h : ͷ�ļ�
//

#pragma once


// CContrastToolDlg �Ի���
class CContrastToolDlg : public CDialogEx
{
// ����
public:
	CContrastToolDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CONTRASTTOOL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


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
