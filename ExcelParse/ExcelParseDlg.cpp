
// ExcelParseDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelParse.h"
#include "ExcelParseDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#include "ExcelHandler.h"
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


// CExcelParseDlg 对话框




CExcelParseDlg::CExcelParseDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcelParseDlg::IDD, pParent)
{
	//m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);	
}

void CExcelParseDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT1, IDC_EDIT_PATH);
	DDX_Control(pDX, IDC_EDIT2, IDC_EDIT2_DEBUG);
}

BEGIN_MESSAGE_MAP(CExcelParseDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(OPEN, &CExcelParseDlg::OnBnClickedOpen)
	ON_BN_CLICKED(REMOVE, &CExcelParseDlg::OnBnClickedRemove)
END_MESSAGE_MAP()


// CExcelParseDlg 消息处理程序

BOOL CExcelParseDlg::OnInitDialog()
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
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);	
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	CString NoteString;
	CExcelParseDlg *DlgOperate;
	NoteString = _T("=======操作步骤==============\r\n");
	NoteString += _T("1.打开Excel文件\r\n");
	NoteString += _T("2.等待单据输出完成\r\n");
	NoteString += _T("3.在OutputFile目录下获取单据\r\n");
	NoteString += _T("=======其他说明==============\r\n");
	NoteString += _T("1.请不要随意更改模版和Excel的格式\r\n");
	NoteString += _T("2.生成单据之后可以使用Word批量打印工具进行打印\r\n");
	NoteString += _T("3.联系开发者点击左上角图标");
	DlgOperate = (CExcelParseDlg *)AfxGetMainWnd();//AfxGetMainWnd();返回CWnd主窗口
	DlgOperate->DebugInfo(NoteString);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CExcelParseDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CExcelParseDlg::OnPaint()
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
HCURSOR CExcelParseDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CExcelParseDlg::DebugInfo(CString String)
{
	SetDlgItemText(IDC_EDIT2, String); 

	CEdit* pedit = (CEdit*)GetDlgItem(IDC_EDIT2);		//For auto turn to down .
	pedit->LineScroll(pedit->GetLineCount());

	UpdateWindow();
}

void CExcelParseDlg::OnBnClickedOpen()
{
	CExcelHandler RealExcelHandler;
	//   TCHAR szFilter[] = _T("文本文件(*.txt)|*.txt|所有文件(*.*)|*.*||");  
	//	注意;代表两个过滤器并行,结尾是双||
	//  注释1|过滤器1.2;过滤器1.2|注释2|过滤器2||
	TCHAR szFilter[] = _T("excel(*.xls;*.xlsx)|*.xls;*.xlsx||");
    // 构造打开文件对话框   
    CFileDialog fileDlg(TRUE, _T("xls"), NULL, 0, szFilter, this);   

    // 显示打开文件对话框   
    if (IDOK == fileDlg.DoModal())   
    {   
        // 如果点击了文件对话框上的“打开”按钮，则将选择的文件路径显示到编辑框里   
        strFilePath_excel = fileDlg.GetPathName();   
        SetDlgItemText(IDC_EDIT1, strFilePath_excel);  
		/*Get size*/
		CFile file_excel(strFilePath_excel,CFile::modeRead);
		UpdateWindow();
		file_excel.Close();

		RealExcelHandler.Excel_AllHandler(strFilePath_excel);

		AfxMessageBox(_T("单据创建成功！"));
    }		
}


void CExcelParseDlg::OnBnClickedRemove()
{
	CExcelHandler RealExcelHandler;
	RealExcelHandler.RemoveDocxFile();

	AfxMessageBox(_T("移除所有单据完成！"));
}
