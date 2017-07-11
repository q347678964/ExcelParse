#include "stdafx.h"
#include "FormatTransfer.h"

/*
params: 0~9,a~f,A~F
return:0~15
*/
char Format_Trans::GetNumFromASCII(char data)
{
	if(data>='0' && data<='9')
		return (data-'0');
	else if(data>='a' && data<='f')
		return (data-'a'+10);
	else if(data>='A' && data<='F')
		return (data-'A'+10);
	else
		return 0;
}
/*
params: unsigned int data,0~0xffffffff
return: src = String,"ffffff"
*/
void Format_Trans::GetASCIIFromNum(unsigned int data,unsigned char *src)
{
	src[0] = Format_Trans::GetHexASCIIFromInt(((data&0xff000000)>>24)/16);
	src[1] = Format_Trans::GetHexASCIIFromInt(((data&0xff000000)>>24)%16);
	src[2] = Format_Trans::GetHexASCIIFromInt(((data&0x00ff0000)>>16)/16);
	src[3] = Format_Trans::GetHexASCIIFromInt(((data&0x00ff0000)>>16)%16);
	src[4] = Format_Trans::GetHexASCIIFromInt(((data&0x0000ff00)>>8)/16);
	src[5] = Format_Trans::GetHexASCIIFromInt(((data&0x0000ff00)>>8)%16);
	src[6] = Format_Trans::GetHexASCIIFromInt(((data&0x000000ff)>>0)/16);
	src[7] = Format_Trans::GetHexASCIIFromInt(((data&0x000000ff)>>0)%16);
}

/*
params: char data,0~9,a~f,A~F
return: 0: No Hex Format , 1:Hex Format
*/
char Format_Trans::IsHexFormat(char data)
{
	if(data>='0' && data<='9')
		return 1;
	else if(data>='a' && data<='f')
		return 1;
	else if(data>='A' && data<='F')
		return 1;
	else
		return 0;
}
/*
params: 0~15
return: 0~9,A~F
*/
char Format_Trans::GetHexASCIIFromInt(char data)
{
	char Index[] = {"0123456789ABCDEF"};
	return Index[data];
}
/*
params: src is hex,255
return: des is hex string "ff"
*/
void Format_Trans::HexToASCII(unsigned char *src,unsigned char*des,unsigned int size)
{
	unsigned int i = 0;
	for(i=0;i<size;i++){
		des[2*i] = Format_Trans::GetHexASCIIFromInt(src[i]/16);
		des[2*i+1] = Format_Trans::GetHexASCIIFromInt(src[i]%16);
	}
}

unsigned long long Format_Trans::GetAddFromAddchar(char *data)
{
	unsigned long long rts = 0;

	rts = GetNumFromASCII(data[0])*(1<<28) + GetNumFromASCII(data[1])*(1<<24) +\
		  GetNumFromASCII(data[2])*(1<<20) + GetNumFromASCII(data[3])*(1<<16) +\
		  GetNumFromASCII(data[4])*(1<<12) + GetNumFromASCII(data[5])*(1<<8) +\
		  GetNumFromASCII(data[6])*(1<<4) + GetNumFromASCII(data[7])*(1<<0);

	return rts;
}

unsigned long long Format_Trans::GetU32FromAddr(unsigned char *data)
{
	unsigned long long rts = 0;

	rts = (data[0])*(1<<24) + (data[1])*(1<<16) +\
		  (data[2])*(1<<8) + (data[3])*(1<<0);

	return rts;
}

char Format_Trans::GotBigWriteFromLittle(char data)
{
	if(data>='a'&&data<='f')
		return 'A'+data-'a';
	else 
		return data;
}


unsigned char Format_Trans::StringCmp(unsigned char *a,char*b,unsigned int length)
{
	unsigned int i = 0;
	for(i=0;i<length;i++)
	{
		if(a[i]!=b[i])
			return 0;
	}
	return 1;
}

CString Format_Trans::ITCS(int i)
{
	CString str;
	str.Format(_T("%d"),i);
	return str;
}

CString Format_Trans::DTCS(double i)
{
	CString str;
	str.Format(_T("%.2f"),i);
	return str;
}

CString Format_Trans::GetDateString(void)
{
	SYSTEMTIME st;
	CString strDate;
	GetLocalTime(&st);
	strDate.Format(_T("%4d-%02d-%2d"),st.wYear,st.wMonth,st.wDay);
	return strDate;
}

CString Format_Trans::GetCurMonthString(void)
{
	SYSTEMTIME st;
	CString strCurMonth;
	GetLocalTime(&st);
	strCurMonth.Format(_T("%2d 月份"),st.wMonth);
	return strCurMonth;
}

CString Format_Trans::GetLastMonthString(void)
{
	SYSTEMTIME st;
	CString strLastMonth;
	GetLocalTime(&st);
	strLastMonth.Format(_T("%2d 月份"),st.wMonth-1);
	return strLastMonth;
}

CString Format_Trans::GetTimeString(void)
{
	SYSTEMTIME st;
	CString strTime;
	GetLocalTime(&st);
	strTime.Format(_T("%2d:%2d:%2d"),st.wHour,st.wMinute,st.wSecond);
	return strTime;
}