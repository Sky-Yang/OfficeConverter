#ifndef _WORD_CONVERTER_H_
#define _WORD_CONVERTER_H_

#include "office/office_converter.h"

class WordConverter: public OfficeConverter
{
public:
    WordConverter();
    virtual ~WordConverter();

    // convert office file to picture
    // if |width| larger than 0, scale the picture by width
    // if |width| is smaller than 0 and |height| larger than 0, scale the picture by height
    // if |width| and |height| is both smaller than 0, convert the picture by it's original size
    virtual bool Convert(const std::wstring& file_path,
                         const std::wstring& output_path,
                         int width, int height) override;
};


#endif // WORD_CONVERTER_H_