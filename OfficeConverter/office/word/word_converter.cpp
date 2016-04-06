#include "StdAfx.h"
#include "office/word/word_converter.h"

#include <sstream>
#include "office/word/word_interfaces.h"

using namespace word;

WordConverter::WordConverter()
{

}

WordConverter::~WordConverter()
{

}

bool WordConverter::Convert(const std::wstring& file_path, 
                            const std::wstring& output_path)
{
    CApplication WordApp;   // WORD����
    WordApp.m_bAutoRelease = true;
    try
    {
        if (CoInitialize(NULL) != S_OK)
        {
            AfxMessageBox(L"��ʼ��COMʱ���ִ���");
            return false;
        }
        if (!WordApp.CreateDispatch(L"Word.Application", NULL))
        {
            AfxMessageBox(L"�޷�����Word����!���Ȱ�װOffice Word!");
            CoUninitialize();
            return false;
        }
    }
    catch (...)
    {
        assert(false && L"��ʼ��ʱ���ִ���");
        return false;
    }
    CString version = WordApp.get_Version();
    int ver = 15;
    try
    {
        ver = _wtoi(version.GetBuffer());
        version.ReleaseBuffer();
    }
    catch (...)
    {
        assert(false && L"ת���汾��ʧ�ܣ������½ӿ�ִ��");
    }

    COleVariant  varfilepath(file_path.c_str());
    COleVariant  varstrnull(L"");
    COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
    COleVariant  vartrue((short)TRUE);
    COleVariant  varfalse((short)FALSE);
    COleVariant  var_file_format((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

    CDocuments docs;        // WORD������������ĵ�
    LPDISPATCH lpDisp;
    docs.AttachDispatch(WordApp.get_Documents());
    try
    {
        switch (ver)
        {
        case OFFICE_97:
            lpDisp = docs.OpenOld(&varfilepath, &varfalse, &vartrue, &varfalse,
                                  &covOptional, &covOptional, &varfalse,
                                  &covOptional, &covOptional, &var_file_format);
            break;
        case OFFICE_2000:
            lpDisp = docs.Open2000(&varfilepath, &varfalse, &vartrue, &varfalse,
                                   &covOptional, &covOptional, &varfalse,
                                   &covOptional, &covOptional, &var_file_format,
                                   &covOptional, &vartrue);
            break;
        case OFFICE_2002:
            lpDisp = docs.Open2002(&varfilepath, &varfalse, &vartrue, &varfalse,
                                   &covOptional, &covOptional, &varfalse,
                                   &covOptional, &covOptional, &var_file_format,
                                   &covOptional, &vartrue, &covOptional,
                                   &covOptional, &covOptional);
            break;
        case OFFICE_2003:
        case OFFICE_2007:
        case OFFICE_2010:
        case OFFICE_2013:
        default:
            lpDisp = docs.Open(&varfilepath, &varfalse, &vartrue, &varfalse,
                               &covOptional, &covOptional, &varfalse, 
                               &covOptional, &covOptional, &var_file_format, 
                               &covOptional, &vartrue, &covOptional, 
                               &covOptional, &covOptional, &covOptional);
        	break;
        }
    }
    catch (...)
    {
        Sleep(600);
        WordApp.Quit(varfalse, covOptional, covOptional);
        assert(false && L"��wordʱ���ִ���");
        return false;
    }

    word::CDocument doc;    // �ĵ�
    CSelection selection;   // ����word�ṩ��ѡ�����
    CRange rng;
    doc.AttachDispatch(lpDisp);
    selection.AttachDispatch(WordApp.get_Selection());
    selection.WholeStory();
    try 
    {
        rng.AttachDispatch(doc.Range(COleVariant(long(0)), COleVariant(selection.get_End())));
    }
    catch (...)
    {
        Sleep(600);
        WordApp.Quit(varfalse, covOptional, covOptional);
        assert(false && L"����wordʱ���ִ���");
        return false;
    }

    long endDoc = rng.get_End();
    long start = 0;
    long end = 0;

    long page_count = rng.ComputeStatistics(2);   //ҳ��
    selection.SetRange(start, end);

    int count = page_count;
    for (int num = 1; num <= count; num++)
    {
        if (num > 1)
        {
            rng = rng.GoToNext(1);
            start = rng.get_Start();
        }

        if (num != page_count)
        {
            rng = rng.GoToNext(1);
            end = rng.get_End();
        }
        else
            end = endDoc;

        selection.SetRange(start, end);
        rng = selection.get_Range();
        try
        {
            rng.CopyAsPicture();
        }
        catch (...)
        {
            WordApp.Quit(varfalse, covOptional, covOptional);
            assert(false && L"ת��ͼƬ���̳��ִ���");
            return false;
        }

        std::wostringstream out_stream;
        out_stream << output_path.c_str() << L"_word_" << num << L".png";
        bool result = Save(out_stream.str());
        if (!result)
        {
            rng.ReleaseDispatch();
            selection.ReleaseDispatch();
            docs.ReleaseDispatch();
            doc.ReleaseDispatch();
            WordApp.Quit(varfalse, covOptional, covOptional);
            WordApp.ReleaseDispatch();
            CoUninitialize();
            return false;
        }
    }

    rng.ReleaseDispatch();
    selection.ReleaseDispatch();
    docs.ReleaseDispatch();
    doc.ReleaseDispatch();
    WordApp.Quit(varfalse, covOptional, covOptional);
    WordApp.ReleaseDispatch();
    CoUninitialize();
    return true;
}
