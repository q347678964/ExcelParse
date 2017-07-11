#include "stdafx.h"
#include "ExcelHandler.h"
#include "ExcelLib/OperateExcelFile.h"
#include "WordLib/OperateWordFile.h"
#include "resource.h"		// ������
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

	BOOL bWorking = Finder.FindFile(strPathName+_T("OutputFile\\��ס¥�վ�\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\��ס¥�վ�\\")+FileName;
		if(CFile::GetStatus(FileName,Fstatus,NULL)){
			CFile::Remove(FileName);
		}
	}

	bWorking = Finder.FindFile(strPathName+_T("OutputFile\\��ס¥֪ͨ\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\��ס¥֪ͨ\\")+FileName;
		if(CFile::GetStatus(FileName,Fstatus,NULL)){
			CFile::Remove(FileName);
		}
	}

	bWorking = Finder.FindFile(strPathName+_T("OutputFile\\д��¥�վ�\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\д��¥�վ�\\")+FileName;
		if(CFile::GetStatus(FileName,Fstatus,NULL)){
			CFile::Remove(FileName);
		}
	}

	bWorking = Finder.FindFile(strPathName+_T("OutputFile\\д��¥֪ͨ\\*.doc"));
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();
		FileName = strPathName+_T("OutputFile\\д��¥֪ͨ\\")+FileName;
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
	DebugInfoString += (CString)"�����ļ���Ϣ: ";
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
	DlgOperate = (CExcelParseDlg *)AfxGetMainWnd();//AfxGetMainWnd();����CWnd������
	DlgOperate->DebugInfo(DebugInfoString);
}

/*
���·��
*/
void CExcelHandler::Excel_ExcelHandler(CString FilePath)
{
	int SheetNumber = 0,SheetColumn = 0,CurColumn = 0,SheetRow = 0,CurRow = 0,i = 0;
	int CurSheet = 0;  //Sheet1����ס¥,Sheet2��д��¥
	int CurConfig = 0;
	int FileCreateCounter = 0;
	double DoubleData = 0.0;
	CString SheetName;
	CString CurPointStringData;

	ExcelOperate.InitExcel();
	ExcelOperate.OpenExcelFile(FilePath);

	SheetNumber = ExcelOperate.GetSheetCount();

	DebugInfoString += (CString)"����� = ";
	DebugInfoString += CExcelHandler::ITCS(SheetNumber);
	CExcelHandler::DebugUpdate();

	while(CurSheet<SheetNumber){				//one excel sheet = two word files
		CurSheet++;				//Frist  to be 1
		CurConfig++;			//First to be 1
		SheetName = ExcelOperate.GetSheetName(CurSheet);
		DebugInfoString += (CString)"���";
		DebugInfoString += SheetName;
		//CExcelHandler::DebugUpdate();
		{			//For Debug
			ExcelOperate.LoadSheet(CurSheet,1);

			SheetRow = ExcelOperate.GetRowCount();	//��
			DebugInfoString += (CString)",���� = ";
			DebugInfoString += CExcelHandler::ITCS(SheetRow);
			//CExcelHandler::DebugUpdate();


			SheetColumn = ExcelOperate.GetColumnCount();	//��
			DebugInfoString += (CString)",���� = ";
			DebugInfoString += CExcelHandler::ITCS(SheetColumn);
			CExcelHandler::DebugUpdate();
		}
		/*begin handler sheet data*/
		ExcelOperate.LoadSheet(CurSheet,1);		//����Sheet1���
		SheetRow = ExcelOperate.GetRowCount();	//��
		SheetColumn = ExcelOperate.GetColumnCount();	//��

		WordOperate.ClearCounter();	//����Word�ļ�����Ϊ0�����漰���ļ�����
		FileCreateCounter = 0;//���ô�ӡ����Ϊ0����ֻ��Ƶ���ӡLog
		/*֪ͨ��************************************************************************/
		CExcelHandler::Excel_ReadConfig(CurConfig);				//1.��ȡ֪ͨ�����ļ����ٶ�ȡ�վ������ļ�
		for(int j=1;j<=SheetRow;j++){				//j������
			if(ExcelOperate.GetCellString(j,CExcelHandler::GetCreateFlagNum()) == "1"){	//�������Ƿ�Ҫ�����ļ�,1�������ļ�
				FileCreateCounter++;
				DebugInfoString += CExcelHandler::ITCS(FileCreateCounter);
				DebugInfoString += "|";
				for(unsigned int i=0;i<CExcelHandler::GetItemNumber();i++){//����config.bin�ļ�����������ȡһ���û�����Ӧ���ݴ���CustomerInfoString
					CurColumn = CExcelHandler::GetExcelInputNumber(i);
					CustomerInfoOutputLabelOrder[i] = CExcelHandler::GetWordOutputNumber(i);	//���20Ŀǰ
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

				WordOperate.WordCreate(CustomerInfoString,CustomerInfoOutputLabelOrder,CExcelHandler::GetItemNumber(),CurConfig);							//3.����Word���ݣ������ݰ�˳��д���ǩ
			}
		}

		CurConfig++;	//��һ�����þ����վ�
		WordOperate.ClearCounter();	//����Word�ļ�����Ϊ0�����漰���ļ�����
		FileCreateCounter = 0;//���ô�ӡ����Ϊ0����ֻ��Ƶ���ӡLog

		/*�վݵ�************************************************************************/
		CExcelHandler::Excel_ReadConfig(CurConfig);				//2.�ٶ�ȡ�վ������ļ�
		for(int j=1;j<=SheetRow;j++){				//j������
			if(ExcelOperate.GetCellString(j,CExcelHandler::GetCreateFlagNum()) == "1"){	//�������Ƿ�Ҫ�����ļ�,1�������ļ�
				FileCreateCounter++;
				DebugInfoString += CExcelHandler::ITCS(FileCreateCounter);
				DebugInfoString += "|";
				for(unsigned int i=0;i<CExcelHandler::GetItemNumber();i++){//����config.bin�ļ�����������ȡһ���û�����Ӧ���ݴ���CustomerInfoString
					CurColumn = CExcelHandler::GetExcelInputNumber(i);
					CustomerInfoOutputLabelOrder[i] = CExcelHandler::GetWordOutputNumber(i);	//���20Ŀǰ
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

				WordOperate.WordCreate(CustomerInfoString,CustomerInfoOutputLabelOrder,CExcelHandler::GetItemNumber(),CurConfig);							//3.����Word���ݣ������ݰ�˳��д���ǩ
			}
		}

	}
	ExcelOperate.CloseExcelFile();
	ExcelOperate.ReleaseExcel();

}
void CExcelHandler::Excel_AllHandler(CString FilePath)
{
	CExcelHandler::RemoveDocxFile();//�Ƴ�*.docx

	CExcelHandler::Excel_ExcelHandler(FilePath);	//����Excel����
}
