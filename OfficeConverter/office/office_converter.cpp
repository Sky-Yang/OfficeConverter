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

bool OfficeConverter::Save(const std::wstring& output_file_path,
                           int width, int height)
{
    // Initialize GDI+.
    GdiplusStartupInput gdiplusStartupInput;
    ULONG_PTR gdiplusToken;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    bool result = false;
    do 
    {
        if (!::OpenClipboard(NULL)) // 打开剪贴板
        {
            assert(false && L"打开剪贴板失败");
            break;
        }

        //////////////////////////////////////////////////////////////////////////
        HENHMETAFILE hEnhMetaFile = (HENHMETAFILE)GetClipboardData(CF_ENHMETAFILE); // 获取剪贴板数据句柄 
        if (hEnhMetaFile == NULL)
        {
            int err = GetLastError();
            assert(false && L"操作剪贴板时出现错误");
            break;
        }

        Gdiplus::Metafile metaFile(hEnhMetaFile);
        int result_widht = width;
        int result_height = height;
        if (width > 0)
        {
            float scale = (static_cast<float>(width))/ metaFile.GetWidth();
            result_height = static_cast<int>(scale * metaFile.GetHeight());
        }
        else if (height > 0)
        {
            float scale = (static_cast<float>(height)) / metaFile.GetHeight();
            result_widht = static_cast<int>(scale * metaFile.GetWidth());
        }
        else
        {
            result_widht = metaFile.GetWidth();
            result_height = metaFile.GetHeight();
        }

        Gdiplus::Bitmap bitmap(result_widht, result_height, PixelFormat24bppRGB);
        Gdiplus::Graphics graphics(&bitmap);
        graphics.Clear(Gdiplus::Color(255, 255, 255));
        Gdiplus::Rect rect(0,0,result_widht, result_height);
        ImageAttributes imAtt;
        imAtt.SetWrapMode(WrapModeTileFlipXY);
        graphics.SetInterpolationMode(InterpolationModeHighQuality);
        graphics.SetPixelOffsetMode(PixelOffsetModeHighQuality);
        graphics.DrawImage(&metaFile, rect, 
                           0, 0, metaFile.GetWidth(), metaFile.GetHeight(), 
                           Gdiplus::UnitPixel, &imAtt);

        CLSID encoderClsid;
        Status stat;

        // Get the CLSID of the PNG encoder.
        if (-1 == GetEncoderClsid(L"image/png", &encoderClsid))
        {
            assert(false && L"获取CLSID失败");
            break;
        }

        //Gdiplus::EncoderParameters parameters;
        //parameters.Count = 1;
        //parameters.Parameter[0].Guid = Gdiplus::EncoderQuality;
        //parameters.Parameter[0].Type = Gdiplus::EncoderParameterValueTypeLong;
        //parameters.Parameter[0].NumberOfValues = 1;
        //int quality = 100;
        //parameters.Parameter[0].Value = &quality;
        //stat = metaFile.Save(output_file_path.c_str(), &encoderClsid, &parameters);
        stat = bitmap.Save(output_file_path.c_str(), &encoderClsid, NULL);
        if (stat != Ok)
        {
            assert(false && L"保存文件夹时出现错误");
        }
        DeleteEnhMetaFile(hEnhMetaFile);

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

        result = true;
    } while (false);

    GdiplusShutdown(gdiplusToken);

    EmptyClipboard();
    CloseClipboard();
    return result;
}
