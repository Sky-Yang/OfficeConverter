#ifndef _EXCEL_CONVERTER_H_
#define _EXCEL_CONVERTER_H_

#include "office/office_converter.h"

class ExcelConverter: public OfficeConverter
{
public:
    ExcelConverter();
    virtual ~ExcelConverter();

    // convert office file to picture
    virtual bool Convert(const std::wstring& file_path,
                         const std::wstring& output_path) override;
};


#endif // _EXCEL_CONVERTER_H_