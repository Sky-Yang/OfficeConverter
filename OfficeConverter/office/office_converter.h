// office_converter.h : office file converter
#pragma once

#include <cassert>
#include <string>

class OfficeConverter
{
public:
    virtual ~OfficeConverter();

    // convert office file to picture
    virtual bool Convert(const std::wstring& file_path,
                         const std::wstring& output_path,
                         int width, int height) = 0;

protected:
    enum VERSION
    {
        OFFICE_97   = 8,
        OFFICE_2000 = 9,
        OFFICE_2002 = 10,
        OFFICE_2003 = 11,
        OFFICE_2007 = 12,
        OFFICE_2010 = 14,
        OFFICE_2013 = 15,
    };

    enum SCALE
    {
        FIT_BY_WIDTH  = 0,
        FIT_BY_HEIGHT = 1,
        FIT_AUTO = 2,
    };

    bool Save(const std::wstring& output_file_path, int width, int height,
              SCALE type);
};