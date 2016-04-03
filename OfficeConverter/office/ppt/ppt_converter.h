#ifndef _PPT_CONVERTER_H_
#define _PPT_CONVERTER_H_

#include "office/office_converter.h"

class PptConverter: public OfficeConverter
{
public:
    PptConverter();
    virtual ~PptConverter();

    // convert office file to picture with |width| and |height|
    virtual bool Convert(const std::wstring& file_path,
                         const std::wstring& output_path,
                         int width, int height) override;
};


#endif // _PPT_CONVERTER_H_