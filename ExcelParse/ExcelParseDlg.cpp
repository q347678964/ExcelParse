
// ExcelParseDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ExcelParse.h"
#include "ExcelParseDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#include "ExcelHandler.h"
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


// CExcelParseDlg �Ի���




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


// CExcelParseDlg ��Ϣ�������

BOOL CExcelParseDlg::OnInitDialog()
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
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);	
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	CString NoteString;
	CExcelParseDlg *DlgOperate;
	NoteString = _T("=======��������==============\r\n");
	NoteString += _T("1.��Excel�ļ�\r\n");
	NoteString += _T("2.�ȴ�����������\r\n");
	NoteString += _T("3.��OutputFileĿ¼�»�ȡ����\r\n");
	NoteString += _T("=======����˵��==============\r\n");
	NoteString += _T("1.�벻Ҫ�������ģ���Excel�ĸ�ʽ\r\n");
	NoteString += _T("2.���ɵ���֮�����ʹ��Word������ӡ���߽��д�ӡ\r\n");
	NoteString += _T("3.��ϵ�����ߵ�����Ͻ�ͼ��");
	DlgOperate = (CExcelParseDlg *)AfxGetMainWnd();//AfxGetMainWnd();����CWnd������
	DlgOperate->DebugInfo(NoteString);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CExcelParseDlg::OnPaint()
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
	//   TCHAR szFilter[] = _T("�ı��ļ�(*.txt)|*.txt|�����ļ�(*.*)|*.*||");  
	//	ע��;������������������,��β��˫||
	//  ע��1|������1.2;������1.2|ע��2|������2||
	TCHAR szFilter[] = _T("excel(*.xls;*.xlsx)|*.xls;*.xlsx||");
    // ������ļ��Ի���   
    CFileDialog fileDlg(TRUE, _T("xls"), NULL, 0, szFilter, this);   

    // ��ʾ���ļ��Ի���   
    if (IDOK == fileDlg.DoModal())   
    {   
        // ���������ļ��Ի����ϵġ��򿪡���ť����ѡ����ļ�·����ʾ���༭����   
        strFilePath_excel = fileDlg.GetPathName();   
        SetDlgItemText(IDC_EDIT1, strFilePath_excel);  
		/*Get size*/
		CFile file_excel(strFilePath_excel,CFile::modeRead);
		UpdateWindow();
		file_excel.Close();

		RealExcelHandler.Excel_AllHandler(strFilePath_excel);

		AfxMessageBox(_T("���ݴ����ɹ���"));
    }		
}


void CExcelParseDlg::OnBnClickedRemove()
{
	CExcelHandler RealExcelHandler;
	RealExcelHandler.RemoveDocxFile();

	AfxMessageBox(_T("�Ƴ����е�����ɣ�"));
}
