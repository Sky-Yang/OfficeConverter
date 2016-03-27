// office_converter.h : office file converter
#pragma once

#include <string>

class OfficeConverter
{
public:
    virtual ~OfficeConverter();

    // convert office file to picture
    virtual bool Convert(const std::wstring& file_path,
                         const std::wstring& output_path) = 0;

protected:
    bool Save(const std::wstring& output_file_path);
};