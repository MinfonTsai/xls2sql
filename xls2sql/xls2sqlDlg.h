// xls2sqlDlg.h : ���Y��
//

#pragma once
#include "afxwin.h"


// Cxls2sqlDlg ��ܤ��
class Cxls2sqlDlg : public CDialog
{
// �غc
public:
	Cxls2sqlDlg(CWnd* pParent = NULL);	// �зǫغc�禡

// ��ܤ�����
	enum { IDD = IDD_XLS2SQL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �䴩


// �{���X��@
protected:
	HICON m_hIcon;

	// ���ͪ��T�������禡
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedButtonSqltest();
	CEdit m_edit_file1;
	afx_msg void OnBnClickedButtonCheck();
	afx_msg void OnStnClickedStatic1();
	CStatic m_txt1;
	CEdit m_dbname;
	CEdit m_tablename;
};
