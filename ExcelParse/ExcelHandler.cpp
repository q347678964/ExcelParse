#include "stdafx.h"
#include "ExcelHandler.h"
#include "ExcelLib/OperateExcelFile.h"
#include "WordLib/OperateWordFile.h"
#include "resource.h"		// 主符号
#include "ExcelParseDlg.h"

#define DEBUG_LOG 0

CExcelParseDlg *DlgOperate;

OperateWordFile WordOperate;

IllusionExcelFile ExcelOperate;

CExcelHandler::CExcelHandler(void)
{
	DebugInfoString = (CString)("");
}

void CExcelHandler::RemoveDocxFile(void)
{
	CString  strPathName;
	CString FileName;
	CFileFind Finder;
	CFileStatus Fstatus;

	GetModuleFileName(NULL,strPathName.GetBuffer(256),256);
	strPathName.ReleaseBuffer(256);
	int nPos  = strPathName.ReverseFind('\\');
	strPathName = strPathName.Left(nPos + 1);

	BOOL bWorking = Finder.FindFile(strPathName+_T("OutputFile\\商住楼收据\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\商住楼收据\\")+FileName;
		if(CFile::GetStatus(FileName,Fstatus,NULL)){
			CFile::Remove(FileName);
		}
	}

	bWorking = Finder.FindFile(strPathName+_T("OutputFile\\商住楼通知\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\商住楼通知\\")+FileName;
		if(CFile::GetStatus(FileName,Fstatus,NULL)){
			CFile::Remove(FileName);
		}
	}

	bWorking = Finder.FindFile(strPathName+_T("OutputFile\\写字楼收据\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\写字楼收据\\")+FileName;
		if(CFile::GetStatus(FileName,Fstatus,NULL)){
			CFile::Remove(FileName);
		}
	}

	bWorking = Finder.FindFile(strPathName+_T("OutputFile\\写字楼通知\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\写字楼通知\\")+FileName;
		if(CFile::GetStatus(FileName,Fstatus,NULL)){
			CFile::Remove(FileName);
		}
	}
	DebugInfoString = (CString)("");
	CExcelHandler::DebugUpdate();

}

void CExcelHandler::Excel_ReadConfig(int num)
{
	CExcelHandler::ReadFileConfig(num);
#if DEBUG_LOG
	DebugInfoString += (CString)"====================================";
	CExcelHandler::DebugUpdate();
	DebugInfoString += (CString)"配置文件信息: ";
	CExcelHandler::DebugUpdate();
	for(unsigned int i=0;i<CExcelHandler::GetItemNumber();i++){
		DebugInfoString += CExcelHandler::ITCS(CExcelHandler::GetExcelInputNumber(i));
		DebugInfoString += (CString)"_";
		DebugInfoString += CExcelHandler::ITCS(CExcelHandler::GetWordOutputNumber(i));
		CExcelHandler::DebugUpdate();
	}
	DebugInfoString += (CString)"====================================";
	CExcelHandler::DebugUpdate();
#endif
}

void CExcelHandler::DebugUpdate(void)
{
	DebugInfoString+="\r\n";
	DlgOperate = (CExcelParseDlg *)AfxGetMainWnd();//AfxGetMainWnd();返回CWnd主窗口
	DlgOperate->DebugInfo(DebugInfoString);
}

/*
表格路径
*/
void CExcelHandler::Excel_ExcelHandler(CString FilePath)
{
	int SheetNumber = 0,SheetColumn = 0,CurColumn = 0,SheetRow = 0,CurRow = 0,i = 0;
	int CurSheet = 0;  //Sheet1是商住楼,Sheet2是写字楼
	int CurConfig = 0;
	int FileCreateCounter = 0;
	double DoubleData = 0.0;
	CString SheetName;
	CString CurPointStringData;

	ExcelOperate.InitExcel();
	ExcelOperate.OpenExcelFile(FilePath);

	SheetNumber = ExcelOperate.GetSheetCount();

	DebugInfoString += (CString)"表格数 = ";
	DebugInfoString += CExcelHandler::ITCS(SheetNumber);
	CExcelHandler::DebugUpdate();

	while(CurSheet<SheetNumber){				//one excel sheet = two word files
		CurSheet++;				//Frist  to be 1
		CurConfig++;			//First to be 1
		SheetName = ExcelOperate.GetSheetName(CurSheet);
		DebugInfoString += (CString)"表格：";
		DebugInfoString += SheetName;
		//CExcelHandler::DebugUpdate();
		{			//For Debug
			ExcelOperate.LoadSheet(CurSheet,1);

			SheetRow = ExcelOperate.GetRowCount();	//行
			DebugInfoString += (CString)",行数 = ";
			DebugInfoString += CExcelHandler::ITCS(SheetRow);
			//CExcelHandler::DebugUpdate();


			SheetColumn = ExcelOperate.GetColumnCount();	//列
			DebugInfoString += (CString)",列数 = ";
			DebugInfoString += CExcelHandler::ITCS(SheetColumn);
			CExcelHandler::DebugUpdate();
		}
		/*begin handler sheet data*/
		ExcelOperate.LoadSheet(CurSheet,1);		//加载Sheet1表格
		SheetRow = ExcelOperate.GetRowCount();	//行
		SheetColumn = ExcelOperate.GetColumnCount();	//列

		WordOperate.ClearCounter();	//设置Word文件计数为0，这涉及到文件命名
		FileCreateCounter = 0;//设置打印计数为0，这只设计到打印Log
		/*通知单************************************************************************/
		CExcelHandler::Excel_ReadConfig(CurConfig);				//1.读取通知配置文件，再读取收据配置文件
		for(int j=1;j<=SheetRow;j++){				//j代表行
			if(ExcelOperate.GetCellString(j,CExcelHandler::GetCreateFlagNum()) == "1"){	//检查该行是否要创建文件,1代表创建文件
				FileCreateCounter++;
				DebugInfoString += CExcelHandler::ITCS(FileCreateCounter);
				DebugInfoString += "|";
				for(unsigned int i=0;i<CExcelHandler::GetItemNumber();i++){//根据config.bin文件的配置来获取一个用户的相应数据存入CustomerInfoString
					CurColumn = CExcelHandler::GetExcelInputNumber(i);
					CustomerInfoOutputLabelOrder[i] = CExcelHandler::GetWordOutputNumber(i);	//最大20目前
					if(CExcelHandler::GetExcelInputType(i) == 'S'){			//String Type
						CustomerInfoString[i] = ExcelOperate.GetCellString(j,CurColumn);
					}else if(CExcelHandler::GetExcelInputType(i) == 'F'){	//Flaot Type
						DoubleData = ExcelOperate.GetCellDouble(j,CurColumn);
						CustomerInfoString[i] = CExcelHandler::DTCS(DoubleData);
					}
			
					DebugInfoString += CustomerInfoString[i];
					DebugInfoString += "|";
				}
				CExcelHandler::DebugUpdate();

				WordOperate.WordCreate(CustomerInfoString,CustomerInfoOutputLabelOrder,CExcelHandler::GetItemNumber(),CurConfig);							//3.处理Word数据，将数据按顺序写入标签
			}
		}

		CurConfig++;	//下一个配置就是收据
		WordOperate.ClearCounter();	//设置Word文件计数为0，这涉及到文件命名
		FileCreateCounter = 0;//设置打印计数为0，这只设计到打印Log

		/*收据单************************************************************************/
		CExcelHandler::Excel_ReadConfig(CurConfig);				//2.再读取收据配置文件
		for(int j=1;j<=SheetRow;j++){				//j代表行
			if(ExcelOperate.GetCellString(j,CExcelHandler::GetCreateFlagNum()) == "1"){	//检查该行是否要创建文件,1代表创建文件
				FileCreateCounter++;
				DebugInfoString += CExcelHandler::ITCS(FileCreateCounter);
				DebugInfoString += "|";
				for(unsigned int i=0;i<CExcelHandler::GetItemNumber();i++){//根据config.bin文件的配置来获取一个用户的相应数据存入CustomerInfoString
					CurColumn = CExcelHandler::GetExcelInputNumber(i);
					CustomerInfoOutputLabelOrder[i] = CExcelHandler::GetWordOutputNumber(i);	//最大20目前
					if(CExcelHandler::GetExcelInputType(i) == 'S'){			//String Type
						CustomerInfoString[i] = ExcelOperate.GetCellString(j,CurColumn);
					}else if(CExcelHandler::GetExcelInputType(i) == 'F'){	//Flaot Type
						DoubleData = ExcelOperate.GetCellDouble(j,CurColumn);
						CustomerInfoString[i] = CExcelHandler::DTCS(DoubleData);
					}
			
					DebugInfoString += CustomerInfoString[i];
					DebugInfoString += "|";
				}
				CExcelHandler::DebugUpdate();

				WordOperate.WordCreate(CustomerInfoString,CustomerInfoOutputLabelOrder,CExcelHandler::GetItemNumber(),CurConfig);							//3.处理Word数据，将数据按顺序写入标签
			}
		}

	}
	ExcelOperate.CloseExcelFile();
	ExcelOperate.ReleaseExcel();

}
void CExcelHandler::Excel_AllHandler(CString FilePath)
{
	CExcelHandler::RemoveDocxFile();//移除*.docx

	CExcelHandler::Excel_ExcelHandler(FilePath);	//处理Excel数据
}
