#include "StdAfx.h"
#include "office/ppt/ppt_converter.h"
#include "office/ppt/ppt_interfaces.h"

using namespace ppt;

PptConverter::PptConverter()
{

}

PptConverter::~PptConverter()
{

}

bool PptConverter::Convert(const std::wstring& file_path,
                           const std::wstring& output_path,
                           int width, int height)
{
    CApplication ppt_app;
    CPresentations presentations;
    CPresentation presentation;
    CSlide slide;
    CSlides slides;

    try
    {
        if (CoInitialize(NULL) != S_OK)
        {
            DWORD erro = GetLastError();
            AfxMessageBox(L"初始化COM时出现错误");
            return false;
        }
        if (!ppt_app.CreateDispatch(_T("PowerPoint.Application"), NULL))
        {
            AfxMessageBox(L"无法启动PowerPoint程序!请先安装Office PowerPoint!");
            CoUninitialize();
            return false;
        }
    }
    catch (...)
    {
        return false;
    }
    CString version = ppt_app.get_Version();
    int ver = 15;
    try
    {
        ver = _wtoi(version.GetBuffer());
        version.ReleaseBuffer();
    }
    catch (...)
    {
        assert(false && L"转换版本号失败，用最新接口执行");
    }
    try
    {
        LPDISPATCH lpDisp;
        ppt_app.m_bAutoRelease = true;
        presentations.AttachDispatch(ppt_app.get_Presentations());
        switch (ver)
        {
        case OFFICE_97:
        case OFFICE_2000:
        case OFFICE_2002:
        case OFFICE_2003:
            lpDisp = presentations.OpenOld(file_path.c_str(), 1, 0, 0);
            break;
        case OFFICE_2007:
            lpDisp = presentations.Open2007(file_path.c_str(), 1, 0, 0, 0);
        case OFFICE_2010:
        case OFFICE_2013:
        default:
            lpDisp = presentations.Open(file_path.c_str(), 1, 0, 0);
            break;
        }
        presentation.AttachDispatch(lpDisp, TRUE);

        //////////////////////////////////////////////////////////////////////////
        // the first way to exporting ppt to images
        //presentation.Export(output_path.c_str(), L"png", width, height);
        //////////////////////////////////////////////////////////////////////////
        /* another way to convert to image files */         
        slides = presentation.get_Slides();
        int pageCount = slides.get_Count();
        for (int i = 1; i <= pageCount; i++)
        {
            slide = slides.Range(COleVariant((long)i));
            slide.Copy();
            std::wstring filename;
            filename.resize(256);
            bool result = false;
            if (0 < swprintf_s(&filename.front(), 256, L"%s%s%04d%s", output_path.c_str(), L"_ppt_", i, L".png"))
                result = Save(filename, width, height);

            if (!result)
            {
                int err = GetLastError();
                slide.ReleaseDispatch();
                slides.ReleaseDispatch();
                presentation.Close();
                presentation.ReleaseDispatch();
                presentations.ReleaseDispatch();
                ppt_app.Quit();
                ppt_app.ReleaseDispatch();
                CoUninitialize();
                return false;
            }
        }
        slide.ReleaseDispatch();
        slides.ReleaseDispatch();
        //////////////////////////////////////////////////////////////////////////
    }
    catch (...)
    {
        assert(false && L"操作ppt时出现错误");
        presentation.Close();
        ppt_app.Quit();
        presentation.ReleaseDispatch();
        presentations.ReleaseDispatch();
        ppt_app.ReleaseDispatch();
        CoUninitialize();
        return false;
    }
    presentation.Close();
    ppt_app.Quit();
    presentation.ReleaseDispatch();
    presentations.ReleaseDispatch();
    ppt_app.ReleaseDispatch();
    CoUninitialize();
    return true;
}
