
// ExcelParseDlg.h : 头文件
//
#ifndef EXCEL_PARSEDLG_H
#define EXCEL_PARSEDLG_H

#pragma once
#include "afxwin.h"


// CExcelParseDlg 对话框
class CExcelParseDlg : public CDialogEx
{
// 构造
public:
	CExcelParseDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EXCELPARSE_DIALOG };

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
	afx_msg void OnBnClickedOpen();
	CString strFilePath_excel;
	void CExcelParseDlg::ExcelHandler();
	CEdit IDC_EDIT_PATH;
	CEdit IDC_EDIT2_DEBUG;
	void CExcelParseDlg::DebugInfo(CString String);
	afx_msg void OnBnClickedRemove();
};
#endif
