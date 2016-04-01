#ifndef _PPT_CONVERTER_H_
#define _PPT_CONVERTER_H_

#include "office/office_converter.h"

class PptConverter: public OfficeConverter
{
public:
    PptConverter(int width, int height);
    virtual ~PptConverter();

    // convert office file to picture
    virtual bool Convert(const std::wstring& file_path,
                         const std::wstring& output_path) override;

private:
    int width_;
    int height_;
};


#endif // _PPT_CONVERTER_H_