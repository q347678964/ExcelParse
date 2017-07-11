#ifndef EXCEL_HANDLER
#define EXCEL_HANDLER

#include "FormatTransfer.h"
#include "ReadConfig.h"
#include "afxwin.h"

class CExcelHandler:public Format_Trans,public ReadConfig
{
	public:
		CExcelHandler(void);
		void DebugUpdate(void);
		void RemoveDocxFile(void);
		void Excel_ReadConfig(int num);
		void Excel_ExcelHandler(CString FilePath);
		void Excel_AllHandler(CString FilePath);
		CString DebugInfoString;
		CString CustomerInfoString[100];
		unsigned int CustomerInfoOutputLabelOrder[100];
};

#endif
