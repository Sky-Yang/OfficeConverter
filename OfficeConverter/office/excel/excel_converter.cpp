#include "StdAfx.h"
#include "office/excel/excel_converter.h"
#include "office/excel/excel_interfaces.h"

using namespace excel;

ExcelConverter::ExcelConverter()
{

}

ExcelConverter::~ExcelConverter()
{

}

bool ExcelConverter::Convert(const std::wstring& file_path,
                             const std::wstring& output_path,
                             int width, int height)
{
    CApplication app;
    CWorkbooks books;
    CWorkbook book;
    CWorksheets sheets;
    CWorksheet sheet;
    CRange range;
    CRange iCell;
    LPDISPATCH lpDisp;
    COleVariant vResult;
    COleVariant
        covTrue((short)TRUE),
        covFalse((short)FALSE),
        covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

    if (CoInitialize(NULL) != S_OK)
    {
        AfxMessageBox(L"初始化COM时出现错误");
        return false;
    }

    if (!app.CreateDispatch(_T("Excel.Application")))
    {
        AfxMessageBox(_T("无法启动Excel程序!请先安装Office Excel!"));
        CoUninitialize();
        return false;
    }

    CString version = app.get_Version();
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
    app.put_UserControl(TRUE);

    books.AttachDispatch(app.get_Workbooks());
    lpDisp = books.Open((LPCTSTR)file_path.c_str(),
                        covOptional, covOptional, covOptional, covOptional, covOptional,
                        covOptional, covOptional, covOptional, covOptional, covOptional,
                        covOptional, covOptional, covOptional, covOptional);

    book.AttachDispatch(lpDisp);

    sheets.AttachDispatch(book.get_Worksheets());

    long page_count = sheets.get_Count();

    CRange used_range;

    for (int i = 1; i < page_count + 1; i++)
    {
        COleVariant vOpt((long)i);
        sheet = sheets.get_Item(vOpt);
        sheet.Activate();

        used_range.AttachDispatch(sheet.get_UsedRange());
        used_range.Select();

        used_range.CopyPicture(1, 1);

        std::wstring filename;
        filename.resize(256);
        bool result = false;
        if (0 < swprintf_s(&filename.front(), 256, L"%s%s%04d%s", output_path.c_str(), L"_excel_", i, L".png"))
            result = Save(filename, width, height, FIT_BY_WIDTH);

        if (!result)
        {
            used_range.ReleaseDispatch();
            range.ReleaseDispatch();
            book.ReleaseDispatch();
            books.ReleaseDispatch();
            sheet.ReleaseDispatch();
            sheets.ReleaseDispatch();
            app.Quit();
            app.ReleaseDispatch();

            CoUninitialize();
            return false;
        }
    }

    used_range.ReleaseDispatch();
    range.ReleaseDispatch();
    book.ReleaseDispatch();
    books.ReleaseDispatch();
    sheet.ReleaseDispatch();
    sheets.ReleaseDispatch();
    app.Quit();
    app.ReleaseDispatch();

    CoUninitialize();
    return true;
}
