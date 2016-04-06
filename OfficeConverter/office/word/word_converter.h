#ifndef _WORD_CONVERTER_H_
#define _WORD_CONVERTER_H_

#include "office/office_converter.h"

class WordConverter: public OfficeConverter
{
public:
    WordConverter();
    virtual ~WordConverter();

    // convert office file to picture
    virtual bool Convert(const std::wstring& file_path,
                         const std::wstring& output_path,
                         int width, int height) override;
};


#endif // WORD_CONVERTER_H_