
// ExcelParseDlg.h : ͷ�ļ�
//
#ifndef EXCEL_PARSEDLG_H
#define EXCEL_PARSEDLG_H

#pragma once
#include "afxwin.h"


// CExcelParseDlg �Ի���
class CExcelParseDlg : public CDialogEx
{
// ����
public:
	CExcelParseDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EXCELPARSE_DIALOG };

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
	afx_msg void OnBnClickedOpen();
	CString strFilePath_excel;
	void CExcelParseDlg::ExcelHandler();
	CEdit IDC_EDIT_PATH;
	CEdit IDC_EDIT2_DEBUG;
	void CExcelParseDlg::DebugInfo(CString String);
	afx_msg void OnBnClickedRemove();
};
#endif
