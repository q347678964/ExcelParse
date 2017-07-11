#include "StdAfx.h"  
#include "OperateWordFile.h"

CString ModuleName[5] = {(CString)("Keep"),(CString)("InputFile\\��ס¥֪ͨģ��.dot"),(CString)("InputFile\\��ס¥�վ�ģ��.dot"),\
											(CString)("InputFile\\д��¥֪ͨģ��.dot"),(CString)("InputFile\\д��¥�վ�ģ��.dot")};

CString OutputName[5] = {(CString)("Keep"),(CString)("OutputFile\\��ס¥֪ͨ\\"),(CString)("OutputFile\\��ס¥�վ�\\"),\
											(CString)("OutputFile\\д��¥֪ͨ\\"),(CString)("OutputFile\\д��¥�վ�\\")};


CString OutputFileName[5] = {(CString)("Keep"),(CString)("��ס¥֪ͨ"),(CString)("��ס¥�վ�"),\
											(CString)("д��¥֪ͨ"),(CString)("д��¥�վ�")};


CString g_Labal[100] = {\
	_T("S0"),\
	_T("S1"),_T("S2"),_T("S3"),_T("S4"),_T("S5"),_T("S6"),_T("S7"),_T("S8"),_T("S9"),_T("S10"),\
	_T("S11"),_T("S12"),_T("S13"),_T("S14"),_T("S15"),_T("S16"),_T("S17"),_T("S18"),_T("S19"),_T("S20"),\
	_T("S21"),_T("S22"),_T("S23"),_T("S24"),_T("S25"),_T("S26"),_T("S27"),_T("S28"),_T("S29"),_T("S30"),\
	_T("S31"),_T("S32"),_T("S33"),_T("S34"),_T("S35"),_T("S36"),_T("S37"),_T("S38"),_T("S39"),_T("S40"),\
	_T("S41"),_T("S42"),_T("S43"),_T("S44"),_T("S45"),_T("S46"),_T("S47"),_T("S48"),_T("S49"),_T("S50"),\
	_T("S51"),_T("S52"),_T("S53"),_T("S54"),_T("S55"),_T("S56"),_T("S57"),_T("S58"),_T("S59"),_T("S60"),\
	_T("S61"),_T("S62"),_T("S63"),_T("S64"),_T("S65"),_T("S66"),_T("S67"),_T("S68"),_T("S69"),_T("S70"),\
	_T("S71"),_T("S72"),_T("S73"),_T("S74"),_T("S75"),_T("S76"),_T("S77"),_T("S78"),_T("S79"),_T("S80"),\
	_T("S81"),_T("S82"),_T("S83"),_T("S84"),_T("S85"),_T("S86"),_T("S87"),_T("S88"),_T("S89"),_T("S90"),\
	_T("S91"),_T("S92"),_T("S93"),_T("S94"),_T("S95"),_T("S96"),_T("S97"),_T("S98"),_T("S99"),\
};

static int WordFileCounter = 0;

void OperateWordFile::ClearCounter(void)
{
	WordFileCounter = 0;
}

CString OperateWordFile::GetToolPath(void)
{
	CString  strPathName;
	GetModuleFileName(NULL,strPathName.GetBuffer(256),256);
	strPathName.ReleaseBuffer(256);
	int nPos  = strPathName.ReverseFind('\\');
	strPathName = strPathName.Left(nPos + 1);
	//AfxMessageBox(strPathName);

	return strPathName;
}

void OperateWordFile::WordCreate(CString *CusomterInfo,unsigned int *WordOrder,int StringNum,int modenum)
{
	CString ModuleStrPath;
	CFileStatus Fstatus;
	CString strSavePath;

	ModuleStrPath = OperateWordFile::GetToolPath();
	ModuleStrPath += ModuleName[modenum];
	//AfxMessageBox(ModuleStrPath);

	COleVariant    covZero((short)0),covTrue((short)TRUE),covFalse((short)FALSE),covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
	covDocxType((short)0),start_line, end_line,dot(ModuleStrPath);//*.dot For module file of word

	CApplication0 wordApp;
	CDocuments docs;
	CDocument0 docx;
	CBookmarks bookmarks;
	CBookmark0 bookmark;
	CRange0 range;
	CCell cell;
	
	if (!wordApp.CreateDispatch(_T("Word.Application")))
	{
		AfxMessageBox(_T("����û�а�װword��Ʒ��"));
		return;
	}

	wordApp.put_Visible(FALSE);
	CString wordVersion = wordApp.get_Version();
	docs = wordApp.get_Documents();
	docx = docs.Add(dot, covOptional, covOptional, covOptional);
	bookmarks = docx.get_Bookmarks();

	for(int i=0;i<StringNum;i++)	//�����Ϣ��ģ��
	{
		//g_Labal[WordOrder[i]]����config.bin��Sxxyy,yy�Ƕ���WordOrder[i]���Ƕ��٣�g_Labal[]����Stringyy,ֻ����������û��ʹ��0��ǩ
		//WordOrder[0] = 1;��һ����ǩ����string1
		bookmark = bookmarks.Item(&_variant_t(g_Labal[WordOrder[i]]));	
		range = bookmark.get_Range();
		range.put_Text(CusomterInfo[i]);
	}
#if 0
	/*��������д��*/
	bookmark = bookmarks.Item(&_variant_t(_T("����")));
	range = bookmark.get_Range();
	range.put_Text(OperateWordFile::GetDateString());

	bookmark = bookmarks.Item(&_variant_t(_T("ˮ���·�")));
	range = bookmark.get_Range();
	range.put_Text(OperateWordFile::GetLastMonthString());

	bookmark = bookmarks.Item(&_variant_t(_T("����·�")));
	range = bookmark.get_Range();
	range.put_Text(OperateWordFile::GetLastMonthString());

	bookmark = bookmarks.Item(&_variant_t(_T("������·�")));
	range = bookmark.get_Range();
	range.put_Text(OperateWordFile::GetLastMonthString());

	bookmark = bookmarks.Item(&_variant_t(_T("�����·�")));
	range = bookmark.get_Range();
	range.put_Text(OperateWordFile::GetCurMonthString());
#endif
	WordFileCounter++;

	/*��ֹû�����ݵ�����*/
	/*
	for(int i=0;i<StringNum;i++){
		if(CusomterInfo[i] == ""){
			CusomterInfo[i] = _T("XXX");
		}
	}*/
	
	strSavePath = OperateWordFile::GetToolPath()+OutputName[modenum]+OperateWordFile::ITCS(WordFileCounter)+OutputFileName[modenum]+_T("_")+CusomterInfo[0]+_T("_")+CusomterInfo[1]+_T(".doc");
	
	/*��ֹ����.��������������*/
	if(CFile::GetStatus(strSavePath,Fstatus,NULL)){
		strSavePath = OperateWordFile::GetToolPath()+OutputName[modenum]+OperateWordFile::ITCS(WordFileCounter)+_T("_")+CusomterInfo[0]+_T("_")+CusomterInfo[1]+_T(".doc");
		WordFileCounter++;
	}

	//AfxMessageBox(strSavePath);
	docx.SaveAs(COleVariant(strSavePath), covOptional, covOptional, covOptional, covOptional,
	covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional);

	// �˳�wordӦ��
	docx.Close(covFalse, covOptional, covOptional);
	wordApp.Quit(covOptional, covOptional, covOptional);
	range.ReleaseDispatch();
	bookmarks.ReleaseDispatch();
	wordApp.ReleaseDispatch();
	//AfxMessageBox(_T("���ɳɹ���"));
}