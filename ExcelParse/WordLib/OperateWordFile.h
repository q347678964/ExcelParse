#pragma once  

#include "CApplication0.h" //word�������   
#include "CDocuments.h" //�ĵ�������   
#include "CDocument0.h" //docx����   
#include "CSelection.h" //��ѡ����   
#include "CCell.h" //������Ԫ��   
#include "CCells.h" //��Ԫ�񼯺�   
#include "CRange0.h" //�ĵ��е�һ��������Χ   
#include "CTable0.h" //�������   
#include "CTables0.h" //��񼯺�   
#include "CRow.h" //������   
#include "CRows.h" //�м���   
#include "CBookmark0.h" //   
#include "CBookmarks.h" //

#include "..\FormatTransfer.h"

class OperateWordFile:public Format_Trans
{
public:
	void ClearCounter(void);
	void WordCreate(CString *CusomterInfo,unsigned int *WordOrder,int StringNum,int modenum);
	CString GetToolPath(void);
};