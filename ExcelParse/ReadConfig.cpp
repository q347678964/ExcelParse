#include "stdafx.h"
#include "ReadConfig.h"

CString g_ConfigFileName[5] = {(CString)("Keep"),(CString)("InputFile\\商住楼通知config.bin"),(CString)("InputFile\\商住楼收据config.bin"),\
												(CString)("InputFile\\写字楼通知config.bin"),(CString)("InputFile\\写字楼收据config.bin")};

ReadConfig::ReadConfig(void)
{
	ItemNumber = 0;
	CreateFlagNum = 0;
	memset(ExcelInputQueue,0,QUEUESIZE);
	memset(WordOutputQueue,0,QUEUESIZE);
}

void ReadConfig::Clean(void)
{
	ItemNumber = 0;
	CreateFlagNum = 0;
	memset(ExcelInputQueue,0,QUEUESIZE);
	memset(WordOutputQueue,0,QUEUESIZE);
}

CString ReadConfig::GetToolPath(void)
{
	CString  strPathName;
	GetModuleFileName(NULL,strPathName.GetBuffer(256),256);
	strPathName.ReleaseBuffer(256);
	int nPos  = strPathName.ReverseFind('\\');
	strPathName = strPathName.Left(nPos + 1);
	//AfxMessageBox(strPathName);

	return strPathName;
}
/*
s0401
s0202
s0303

sxxyy
xx For ExcelInputQueue
yy For WordOutputQueue
*/
void ReadConfig::ReadFileConfig(int num)
{
	unsigned long long File_Length = 0,File_CurRP = 0;
	unsigned char CurReadBuffer[10];
	CString ConfigFilePath = ReadConfig::GetToolPath() + g_ConfigFileName[num];
	CFile pConfigFile(ConfigFilePath,CFile::modeRead);//Open cfg File

	File_Length = pConfigFile.GetLength();

	ReadConfig::Clean();	//Clear last data 

	while(File_CurRP<File_Length-1)
	{
		pConfigFile.Seek(File_CurRP,CFile::begin);
		pConfigFile.Read(CurReadBuffer,1);
		if(CurReadBuffer[0] == 'S' || CurReadBuffer[0] == 'F'){	//以S开头，代表是String类型的Excel数据
			ExcelInputType[ItemNumber] = CurReadBuffer[0];	//    'S'/'F'
			pConfigFile.Read(CurReadBuffer,4);
			if(ItemNumber < QUEUESIZE){
				ExcelInputQueue[ItemNumber] = (CurReadBuffer[0]-'0')*10+(CurReadBuffer[1]-'0');
				WordOutputQueue[ItemNumber] = (CurReadBuffer[2]-'0')*10+(CurReadBuffer[3]-'0');
				ItemNumber++;
				File_CurRP+=4;
			}
		}else if(CurReadBuffer[0] == 'C'){	//以C开头，代表创建文件标志的列号
			pConfigFile.Read(CurReadBuffer,2);
			CreateFlagNum = (CurReadBuffer[0]-'0')*10+(CurReadBuffer[1]-'0');
			File_CurRP+=2;
		}
		File_CurRP++;
	}

	pConfigFile.Close();
}

unsigned int ReadConfig::GetExcelInputNumber(unsigned int num)
{
	if(num < QUEUESIZE){
		return ExcelInputQueue[num];
	}else{
		return 0;
	}
}
unsigned int ReadConfig::GetWordOutputNumber(unsigned int num)
{
	if(num < QUEUESIZE){
		return WordOutputQueue[num];
	}else{
		return 0;
	}
}

char ReadConfig::GetExcelInputType(unsigned int num)
{
	if(num < QUEUESIZE){
		return ExcelInputType[num];
	}else{
		return 0;
	}
}	

unsigned int ReadConfig::GetItemNumber(void)
{
		return ItemNumber;
}	

unsigned int ReadConfig::GetCreateFlagNum(void)
{
		return CreateFlagNum;
}	