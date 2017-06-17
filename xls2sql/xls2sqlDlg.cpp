// xls2sqlDlg.cpp : 實作檔
//

#include "stdafx.h"
#include "BasicExcelVC6.hpp"
#include "xls2sql.h"
#include "xls2sqlDlg.h"
#include "tomysql.h"

#pragma   comment(lib, "ws2_32.lib ")
using namespace YExcel;

//#define DEFAULT_XLSFILE   "example1.xls"
#define  DEFAULT_XLSFILE   "test.xls"
#define  DEFAULT_SQLDBNAME   "test"
#define  DEFAULT_TABLENAME   "class1"
#define  SQL_Ins_Table_record   "insert into %s(name,sid,MAC) values('%s',%d,'%s')"

Cxls2sqlDlg *pDlg;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 對 App About 使用 CAboutDlg 對話方塊

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 對話方塊資料
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支援

// 程式碼實作
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// Cxls2sqlDlg 對話方塊




Cxls2sqlDlg::Cxls2sqlDlg(CWnd* pParent /*=NULL*/)
	: CDialog(Cxls2sqlDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void Cxls2sqlDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_FILE1, m_edit_file1);
	DDX_Control(pDX, IDC_STATIC1, m_txt1);
	DDX_Control(pDX, IDC_EDIT_DBNAME, m_dbname);
	DDX_Control(pDX, IDC_EDIT_TABLENAME, m_tablename);
}

BEGIN_MESSAGE_MAP(Cxls2sqlDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDOK, &Cxls2sqlDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON_SQLTEST, &Cxls2sqlDlg::OnBnClickedButtonSqltest)
	ON_BN_CLICKED(IDC_BUTTON_CHECK, &Cxls2sqlDlg::OnBnClickedButtonCheck)
	ON_STN_CLICKED(IDC_STATIC1, &Cxls2sqlDlg::OnStnClickedStatic1)
END_MESSAGE_MAP()


// Cxls2sqlDlg 訊息處理常式

BOOL Cxls2sqlDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 將 [關於...] 功能表加入系統功能表。

	// IDM_ABOUTBOX 必須在系統命令範圍之中。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 設定此對話方塊的圖示。當應用程式的主視窗不是對話方塊時，
	// 框架會自動從事此作業
	SetIcon(m_hIcon, TRUE);			// 設定大圖示
	SetIcon(m_hIcon, FALSE);		// 設定小圖示

	// TODO: 在此加入額外的初始設定
	m_edit_file1.SetWindowText(DEFAULT_XLSFILE);
    pDlg = (Cxls2sqlDlg *)AfxGetMainWnd();
	m_dbname.SetWindowText(DEFAULT_SQLDBNAME);
	m_tablename.SetWindowText(DEFAULT_TABLENAME);

	return TRUE;  // 傳回 TRUE，除非您對控制項設定焦點
}

void Cxls2sqlDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// 如果將最小化按鈕加入您的對話方塊，您需要下列的程式碼，
// 以便繪製圖示。對於使用文件/檢視模式的 MFC 應用程式，
// 框架會自動完成此作業。

void Cxls2sqlDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 繪製的裝置內容

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 將圖示置中於用戶端矩形
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 描繪圖示
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// 當使用者拖曳最小化視窗時，
// 系統呼叫這個功能取得游標顯示。
HCURSOR Cxls2sqlDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

BasicExcel e;
void Cxls2sqlDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	// OnOK();
  
  e.Load("example1.xls");
  BasicExcelWorksheet* s1;

	BasicExcelWorksheet* sheet1 = e.GetWorksheet("Sheet1");

 // if (sheet1)
 // {
 //   size_t maxRows = sheet1->GetTotalRows();
 //   size_t maxCols = sheet1->GetTotalCols();
 // }

if (sheet1)
	{
		size_t maxRows = sheet1->GetTotalRows();
		size_t maxCols = sheet1->GetTotalCols();
		cout << "Dimension of " << sheet1->GetAnsiSheetName() << " (" << maxRows << ", " << maxCols << ")" << endl;

		printf("          ");
		for (size_t c=0; c<maxCols; ++c) printf("%10d", c+1);
		cout << endl;

		for (size_t r=0; r<maxRows; ++r)
		{
			printf("%10d", r+1);
			for (size_t c=0; c<maxCols; ++c)
			{
				BasicExcelCell* cell = sheet1->Cell(r,c);
				switch (cell->Type())
				{
					case BasicExcelCell::UNDEFINED:
						printf("          ");
						break;

					case BasicExcelCell::INT:
						printf("%10d", cell->GetInteger());
						break;

					case BasicExcelCell::DOUBLE:
						printf("%10.6lf", cell->GetDouble());
						break;

					case BasicExcelCell::STRING:
						printf("%10s", cell->GetString());
						break;

					case BasicExcelCell::WSTRING:
						wprintf(L"%10s", cell->GetWString());
						break;
				}
			}
			cout << endl;
		}
	}
	cout << endl;
	e.SaveAs("example2.xls");

	// Create a new workbook with 2 worksheets and write some contents.
	e.New(2);
	e.RenameWorksheet("Sheet1", "Test1");
	BasicExcelWorksheet* sheet = e.GetWorksheet("Test1");
	BasicExcelCell* cell;
	if (sheet)
	{
		for (size_t c=0; c<4; ++c)
		{
			cell = sheet->Cell(0,c);
			cell->Set((int)c);
		}

		cell = sheet->Cell(1,3);
		cell->SetDouble(3.141592654);

		sheet->Cell(1,4)->SetString("Test str1");
		sheet->Cell(2,0)->SetString("Test str2");
		sheet->Cell(2,5)->SetString("Test str1");

		sheet->Cell(4,0)->SetDouble(1.1);
		sheet->Cell(4,1)->SetDouble(2.2);
		sheet->Cell(4,2)->SetDouble(3.3);
		sheet->Cell(4,3)->SetDouble(4.4);
		sheet->Cell(4,4)->SetDouble(5.5);

		sheet->Cell(4,4)->EraseContents();
	}

	sheet = e.AddWorksheet("Test2", 1);
	sheet = e.GetWorksheet(1); 
	if (sheet)
	{
		sheet->Cell(1,1)->SetDouble(1.1);
		sheet->Cell(2,2)->SetDouble(2.2);
		sheet->Cell(3,3)->SetDouble(3.3);
		sheet->Cell(4,4)->SetDouble(4.4);
		sheet->Cell(70,2)->SetDouble(5.5);
	}
	e.SaveAs("example3.xls");

	// Load the newly created sheet and display its contents
	e.Load("example3.xls");

	size_t maxSheets = e.GetTotalWorkSheets();
	cout << "Total number of worksheets: " << e.GetTotalWorkSheets() << endl;
	for (size_t i=0; i<maxSheets; ++i)
	{
		BasicExcelWorksheet* sheet = e.GetWorksheet(i);
		if (sheet)
		{
			size_t maxRows = sheet->GetTotalRows();
			size_t maxCols = sheet->GetTotalCols();
			cout << "Dimension of " << sheet->GetAnsiSheetName() << " (" << maxRows << ", " << maxCols << ")" << endl;

			if (maxRows>0)
			{
				printf("          ");
				for (size_t c=0; c<maxCols; ++c) printf("%10d", c+1);
				cout << endl;
			}

			for (size_t r=0; r<maxRows; ++r)
			{
				printf("%10d", r+1);
				for (size_t c=0; c<maxCols; ++c)
				{
					cout << setw(10) << *(sheet->Cell(r,c));	// Another way of printing a cell content.				
				}
				cout << endl;
			}
			if (i==0)
			{
				ofstream f("example4.csv");
				sheet->Print(f, ',', '\"');	// Save the first sheet as a CSV file.
				f.close();
			}
		}
		cout << endl;
	}

}

ToMysql* schoolsql = new ToMysql;

bool Open_SchoolSQL( void )
{
char dbname[100];

     pDlg->m_dbname.GetWindowText(dbname,99);

   if( schoolsql->ConnMySQL("localhost","3306" ,dbname,"root","0000","utf8","ok")==0 )
	{
		//m_txt1.SetWindowText("Connect MySQL success!"); //   
		//pDlg->MessageBox("連接 MySQL Server: test 成功");
		pDlg->m_txt1.SetWindowText( "連接 MySQL Server: test 成功" );
	}
	else
	{
		pDlg->MessageBox("連接 MySQL  Server 失敗");
		//m_txt1.SetWindowText("FAIL");
		return false;
	}
	char *setname = "set names big5";
	mysql_query(&schoolsql->mysql,setname); 

	return true;

}

void Cxls2sqlDlg::OnBnClickedButtonSqltest()
{
	char Msg[500];
	ToMysql* vspdctomysql = new ToMysql;

	if( vspdctomysql->ConnMySQL("localhost","3306" ,"test","root","0000","utf8","ok")==0 )
	{
		//m_txt1.SetWindowText("Connect MySQL success!"); //
		MessageBox("連接 MySQL Server: test 成功");
	}
	else
		MessageBox("連接 MySQL  Server 失敗");
		//m_txt1.SetWindowText("FAIL");
	
	char *setname = "set names big5";
	mysql_query(&vspdctomysql->mysql,setname); 

	//查詢
	char * SQL = "SELECT name,sid,MAC  FROM class1";
	string str = vspdctomysql->SelectData(SQL,3,Msg);
	if( str.length() > 0 )
	{
		MessageBox(str.data());
		//printf("查詢成功/r/n");
		//m_txt1.SetWindowText("查詢成功!");

		/*
		char * SQL1 = "insert into class1(name,sid,MAC) values('我的',10,'77:88:99')";
		if(vspdctomysql->InsertData(SQL1,Msg)==0)
			m_txt1.SetWindowText("插入一筆記錄成功!");


        char * SQL2 = "update class1 set name='已修改了',sid=20, MAC='11:22:33' where id=3";

		if(vspdctomysql->UpdateData(SQL2,Msg) == 0)
		m_txt1.SetWindowText("更新記錄成功!");

        //刪除
		char *SQL4 = "delete from class1 where id=1";
		if(vspdctomysql->DeleteData(SQL4,Msg) == 0)
			m_txt1.SetWindowText("刪除成功!");
       */
	}

	  vspdctomysql->CloseMySQLConn();	

}

void Cxls2sqlDlg::OnBnClickedButtonCheck()
{
	// TODO: Add your control notification handler code here
	BasicExcel e;
	BasicExcelCell* cell; 
	char filepath[MAX_PATH+1];
	char Msg[200];
	//char idx[1000];
    char row[6];
	wchar_t widx[1000];
    char *sheetname;
    wchar_t *UTF8_sheetname;
	
	wchar_t wcell[100];
    char *bcell;
    int icell;
    float fcell;
	size_t r,c;
	size_t topcell_row,topcell_col, namecell_col;   //要放入xls表格中, 要寫入SQL內容的資料定位點
	char Ins_SQL[200];
	int sid, len;
	char name[100];
	char MACaddr[100];
	char tablename[100];
	char *insstatic="插入一筆內容: %s 到SQL資料庫! ";
	char insmessage[100];
    bool name_found, mac_found;

     m_edit_file1.GetWindowText(filepath,MAX_PATH);

	if( e.Load( filepath ) )
	{
        size_t t1;

        t1 =0;
         e.GetSheetName(t1,widx);  // 取得第一份表格(0) 的名字Sheetname , 以 '雙位元' 形式

		BasicExcelWorksheet* sheet1 = e.GetWorksheet(widx);
	   if( sheet1 )
	   {
         size_t maxRows = sheet1->GetTotalRows();  // Y軸有多少列(由上到下數)
		 size_t maxCols = sheet1->GetTotalCols();  // X軸有多少行(由左到右數)

		 topcell_row=9999;
		 topcell_col=9999;

		 //  ===========  尋找 並且 定位 1   ===========

			for (r=0; r<maxRows; ++r)
			{
				//printf("%10d", r+1);
				//itoa(r+1,row,10);
				//m_txt1.SetWindowText(row);

				for ( c=0; c<maxCols; ++c)
				{
					cell = sheet1->Cell(r,c);
					switch (cell->Type())
					{
					case BasicExcelCell::UNDEFINED:
						//printf("          ");
						break;
					case BasicExcelCell::INT:
						icell = cell->GetInteger();
						//printf("%10d", cell->GetInteger());
						if( icell == 1)
						{
							 topcell_row = r;
							 topcell_col = c;
							 c=maxCols; r=maxRows;
						}
						break;

					case BasicExcelCell::DOUBLE:   // 班級, 座號
						fcell = cell->GetDouble();
						//printf("%10.6lf", cell->GetDouble());
						if( fcell == 1)
						{
							 topcell_row = r;
							 topcell_col = c;
							 c=maxCols; r=maxRows;
						}
						break;	

					case BasicExcelCell::STRING:   //學號
						bcell = (char *)cell->GetString();
						//printf("%10s", cell->GetString());
						if( !strcmp(bcell,"1") || !strcmp(bcell,"01"))
						{
							 topcell_row = r;
							 topcell_col = c;
							 c=maxCols; r=maxRows;
						}
						break;

					case BasicExcelCell::WSTRING:   //姓名
						//wcell = (char *)cell->GetWString();
						swprintf(wcell, L"%s", cell->GetWString());
						break;
					} // swtich
				} // for(size_t c=0; c<maxCols.....
				//	cout << endl;
			} //for (size_t r=0; r<maxRows
              

			//  ===========  尋找 座號 姓名的規格  ===========
			//  A. 從左上角開始 , 有一個格 "內容為 1" , 代表user的 號碼
			//  B. 這內容為1的右邊, 找第一格 "內容為中文", 代表user的 姓名


			if( topcell_row < 9999 )
			{
				for ( c=topcell_col; c<maxCols; ++c)
				{
					cell = sheet1->Cell(topcell_row,c);
					switch (cell->Type())
					{
					  	case BasicExcelCell::STRING:   //ASCII英文名 或 數字 ==> 學號等
						bcell = (char *)cell->GetString();
						//printf("%10s", cell->GetString());
						break;

						case BasicExcelCell::WSTRING:   //姓名
						//wcell = (char *)cell->GetWString();
						swprintf(wcell, L"%s", cell->GetWString());
						namecell_col = c;           //定位姓名
						c=maxCols;
						break;

					} // switch
				} // for (size_t c
			}  //if( topcell_row


		//  ===========  寫入 MySQL( 座號, 姓名, MAC等等)  ===========
         if( Open_SchoolSQL( ) )
		 {
			 //if( maxRows > topcell_row+10 )    // 測試專用: 只允許處理20筆記錄
			//		maxRows = topcell_row+10; 
          // if( topcell_row > 1 )  topcell_row--;

			sid = 0;   strcpy(MACaddr,"");

				for( r=topcell_row; r<maxRows; ++r)
				{
					cell = sheet1->Cell(r,topcell_col);  // ====> 序號
					icell=0;  fcell=0; bcell=0;  name_found=false;  mac_found=false;
					switch (cell->Type())
					{
						case BasicExcelCell::INT:
						icell = cell->GetInteger();
						sid = icell;
						break;

						case BasicExcelCell::DOUBLE:   // 班級, 座號
						fcell = cell->GetDouble();
						sid = (int)fcell;
						break;	

						case BasicExcelCell::STRING:   //學號
						bcell = (char *)cell->GetString();
						sid = atoi(bcell);
						namecell_col = topcell_col-1;
						break;
					}

                   
					cell = sheet1->Cell(r,namecell_col);  // ====> 用戶名
					icell=0;  fcell=0; bcell=0;
					switch (cell->Type())
					{
						case BasicExcelCell::STRING:   //ASCII英文名
						bcell = (char *)cell->GetString();
						strcpy(name,bcell);
						//name_found=true;
						break;
						case BasicExcelCell::WSTRING:   //姓名
						//wcell = (char *)cell->GetWString();
						swprintf(wcell, L"%s", cell->GetWString());
						name_found=true;
						break;
					}
				    if( !name_found )
					{
					    namecell_col = topcell_col+1;

						cell = sheet1->Cell(r,namecell_col);  // ====> 用戶名
						icell=0;  fcell=0; bcell=0;
						switch (cell->Type())
						{
							case BasicExcelCell::STRING:   //ASCII英文名
								bcell = (char *)cell->GetString();
								strcpy(name,bcell);
								//name_found=true;
							break;
							case BasicExcelCell::WSTRING:   //姓名
								//wcell = (char *)cell->GetWString();
								swprintf(wcell, L"%s", cell->GetWString());
							name_found=true;
							break;
						}

					}

				   
					cell = sheet1->Cell(r,namecell_col+1);  // ====> MAC位址 or 學號
					icell=0;  fcell=0; bcell=0;
					switch (cell->Type())
					{
						case 0:
						strcpy(MACaddr,"");
						break;
						case BasicExcelCell::DOUBLE:   // 學號(Float數字)
						fcell = cell->GetDouble();
						sprintf(MACaddr,"%d", cell->GetDouble());
						if( strlen(MACaddr) >= 6 ) mac_found = true;
							break;	
						case BasicExcelCell::STRING:   // 學號(字串)
						bcell = (char *)cell->GetString();
						strcpy(MACaddr,bcell);
							if( strlen(MACaddr) >= 6 )	mac_found = true;
							break;
					}
				   
					if( ! mac_found )
					{

						cell = sheet1->Cell(r,namecell_col+2);  // ====> MAC位址 or 學號
						icell=0;  fcell=0; bcell=0;
						switch (cell->Type())
						{
						 case 0:
						   strcpy(MACaddr,"");
						   break;
						 case BasicExcelCell::DOUBLE:   // 學號(Float數字)
							fcell = cell->GetDouble();
							sprintf(MACaddr,"%d", cell->GetDouble());
							
							break;	
						 case BasicExcelCell::STRING:   // 學號(字串)
							bcell = (char *)cell->GetString();
							strcpy(MACaddr,bcell);
							
							break;
						}

					}
					//  把 雙位元wchar 轉為  單位元 char	
					int nIndex = WideCharToMultiByte(CP_ACP, 0, wcell, -1, NULL, 0, NULL, NULL);
					char *pAnsi = new char[nIndex + 1];
					WideCharToMultiByte(CP_ACP, 0, wcell, -1, pAnsi, nIndex, NULL, NULL);

					 pDlg->m_tablename.GetWindowText(tablename,99);

					sprintf( Ins_SQL,SQL_Ins_Table_record ,tablename,pAnsi,sid,MACaddr);
					sprintf( insmessage, insstatic, pAnsi );
					
					delete pAnsi;
					
 					if(schoolsql->InsertData(Ins_SQL,Msg)==0)
						m_txt1.SetWindowText( insmessage );

				}//for( r=topcel

					m_txt1.SetWindowText( " XLS轉換並寫入SQL 完成");

		 } //if( Open_SchoolSQL
			else
				m_txt1.SetWindowText( "連接 MySQL  Server 失敗");

		} //if( sheet1 )
	    else
			MessageBox("EXCEL檔中找不到 [名單] 表格");

	} //if( e.Load
	else
       MessageBox("EXCEL檔開啟失敗! 請確認 xls存在而且在關閉狀態.");
	
}

void Cxls2sqlDlg::OnStnClickedStatic1()
{
	// TODO: Add your control notification handler code here
}
