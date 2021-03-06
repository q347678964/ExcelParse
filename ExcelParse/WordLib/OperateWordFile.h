#pragma once  

#include "CApplication0.h" //word程序对象   
#include "CDocuments.h" //文档集对象   
#include "CDocument0.h" //docx对象   
#include "CSelection.h" //所选内容   
#include "CCell.h" //单个单元格   
#include "CCells.h" //单元格集合   
#include "CRange0.h" //文档中的一个连续范围   
#include "CTable0.h" //单个表格   
#include "CTables0.h" //表格集合   
#include "CRow.h" //单个行   
#include "CRows.h" //行集合   
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