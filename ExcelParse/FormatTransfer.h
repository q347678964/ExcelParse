#ifndef FORMAT_TRANSFER_H
#define FORMAT_TRANSFER_H


class Format_Trans 
{
	public:
		char IsHexFormat(char data);
		char GetNumFromASCII(char data);
		void GetASCIIFromNum(unsigned int data,unsigned char *src);
		char GetHexASCIIFromInt(char data);
		void HexToASCII(unsigned char *src,unsigned char*des,unsigned int size);
		unsigned long long GetU32FromAddr(unsigned char *data);
		unsigned long long GetAddFromAddchar(char* data);
		char Format_Trans::GotBigWriteFromLittle(char data);
		unsigned char Format_Trans::StringCmp(unsigned char *a,char*b,unsigned int length);
		CString ITCS(int i);
		CString DTCS(double i);
		CString GetDateString(void);
		CString GetTimeString(void);
		CString GetCurMonthString(void);
		CString GetLastMonthString(void);
};
#endif
