#include "StdAfx.h"
#include "office/ppt/ppt_converter.h"

#include <sstream>
#include "office/ppt/ppt_interfaces.h"

using namespace ppt;

PptConverter::PptConverter()
{

}

PptConverter::~PptConverter()
{

}

bool PptConverter::Convert(const std::wstring& file_path,
                            const std::wstring& output_path)
{
    CApplication ppt_app;
    CPresentations presentations;
    CPresentation presentation;
    //CSlide slide;
    //CSlides slides;

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
            AfxMessageBox(L"无法启动PowerPoint程序!");
            CoUninitialize();
            return false;
        }
    }
    catch (...)
    {
        return false;
    }

    ppt_app.m_bAutoRelease = true;
    ppt_app.put_Visible(long(1));
    ppt_app.put_WindowState(long(2));
    presentations.AttachDispatch(ppt_app.get_Presentations());
    presentations.Open(file_path.c_str(), TRUE, 1, 1);
    presentation.AttachDispatch(ppt_app.get_ActivePresentation(), TRUE);    
    presentation.Export(output_path.c_str(), L"png", ppt_app.get_Width(), ppt_app.get_Height());

    //////////////////////////////////////////////////////////////////////////
    /* another way to convert to image files */
    /*
    slides = presentation.get_Slides();
    int pageCount = slides.get_Count();
    for (int i = 1; i <= pageCount; i++)
    {
        slide = slides.Range(COleVariant((long)i));
        slide.Copy();
        std::wostringstream out_stream;
        out_stream << output_path.c_str() << L"_ppt_" << i << L".png";
        bool result = Save(out_stream.str());
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
    */
    //////////////////////////////////////////////////////////////////////////
    presentation.Close();
    ppt_app.Quit();
    presentation.ReleaseDispatch();
    presentations.ReleaseDispatch();
    ppt_app.ReleaseDispatch();
    CoUninitialize();
    return true;
}
