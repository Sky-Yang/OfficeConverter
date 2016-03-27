#include "stdafx.h"
#include "office/office_converter.h"

#include <gdiplus.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace Gdiplus;

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
    UINT num = 0;   // number of image encoders
    UINT size = 0;  // size of the image encoder array in bytes

    ImageCodecInfo* pImageCodecInfo = NULL;
    GetImageEncodersSize(&num, &size);
    if (size == 0)
        return -1;  // Failure

    pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
    if (pImageCodecInfo == NULL)
        return -1;  // Failure

    GetImageEncoders(num, size, pImageCodecInfo);

    for (UINT j = 0; j < num; ++j)
    {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0)
        {
            *pClsid = pImageCodecInfo[j].Clsid;
            free(pImageCodecInfo);
            return j;  // Success
        }
    }

    free(pImageCodecInfo);
    return -1;  // Failure
}

OfficeConverter::~OfficeConverter()
{

}

bool OfficeConverter::Save(const std::wstring& output_file_path)
{
    do 
    {
        if (!::OpenClipboard(NULL)) // 打开剪贴板
        {
            assert(false && L"打开剪贴板失败");
            break;
        }

        HENHMETAFILE hEnhMetaFile = (HENHMETAFILE)GetClipboardData(CF_ENHMETAFILE); // 获取剪贴板数据句柄 
        if (hEnhMetaFile == NULL)
        {
            int err = GetLastError();
            assert(false && L"操作剪贴板时出现错误");
            break;
        }
        // Initialize GDI+.
        GdiplusStartupInput gdiplusStartupInput;
        ULONG_PTR gdiplusToken;
        GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

        Gdiplus::Metafile metaFile(hEnhMetaFile);
        CLSID encoderClsid;
        Status stat;

        // Get the CLSID of the PNG encoder.
        GetEncoderClsid(L"image/png", &encoderClsid);
        stat = metaFile.Save(output_file_path.c_str(), &encoderClsid, NULL);
        if (stat != Ok)
        {
            assert(false && L"保存文件夹时出现错误");
        }
        DeleteEnhMetaFile(hEnhMetaFile);

        GdiplusShutdown(gdiplusToken);

        //////////////////////////////////////////////////////////////////////////
        /*  example of saving as metafile
        BOOL ba = ::IsClipboardFormatAvailable(CF_ENHMETAFILE);
        HENHMETAFILE hEnhMetaFile = NULL;
        hEnhMetaFile = (HENHMETAFILE)GetClipboardData(CF_ENHMETAFILE);
        if (hEnhMetaFile == NULL)
        {
            int err = GetLastError();
            assert(false && L"操作剪贴板时出现错误");
            break;
        }

        HENHMETAFILE hMetaFile = CopyEnhMetaFile(hEnhMetaFile,
                                                 output_file_path.c_str());
        if (hEnhMetaFile == NULL)
        {
            int err = GetLastError();
            assert(false && L"保存文件夹时出现错误");
            break;
        }
        DeleteEnhMetaFile(hEnhMetaFile);
        */
        //////////////////////////////////////////////////////////////////////////

        EmptyClipboard();
        CloseClipboard();
        return true;
    } while (false);

    EmptyClipboard();
    CloseClipboard();
    return false;
}
