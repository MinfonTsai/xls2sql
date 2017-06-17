// xls2sqlDlg.h : 標頭檔
//

#pragma once
#include "afxwin.h"


// Cxls2sqlDlg 對話方塊
class Cxls2sqlDlg : public CDialog
{
// 建構
public:
	Cxls2sqlDlg(CWnd* pParent = NULL);	// 標準建構函式

// 對話方塊資料
	enum { IDD = IDD_XLS2SQL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支援


// 程式碼實作
protected:
	HICON m_hIcon;

	// 產生的訊息對應函式
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
