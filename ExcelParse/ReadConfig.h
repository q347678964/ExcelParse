#pragma once  

#define QUEUESIZE 255

class ReadConfig
{
private:
	unsigned int ItemNumber;
	unsigned int ExcelInputQueue[QUEUESIZE];
	char ExcelInputType[QUEUESIZE];
	unsigned int WordOutputQueue[QUEUESIZE];
	unsigned char CreateFlagNum;
public:
	ReadConfig(void);
	void Clean(void);
	void ReadFileConfig(int num);
	CString GetToolPath(void);
	unsigned int GetExcelInputNumber(unsigned int num);
	unsigned int GetWordOutputNumber(unsigned int num);
	char GetExcelInputType(unsigned int num);
	unsigned int GetItemNumber(void);
	unsigned int GetCreateFlagNum(void);
};