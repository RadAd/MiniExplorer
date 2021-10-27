#pragma once

#include <Windows.h>

#include "Format.h"

#ifdef _DEBUG
#define TRACE _trace
#else
#define TRACE false && _trace
#endif

inline bool _trace(_In_z_ _Printf_format_string_ char const* format, ...)
{
    std::string buffer;
    va_list args;
    va_start(args, format);
    Format(buffer, format, args);
    va_end(args);
    OutputDebugStringA(buffer.c_str());
    return true;
}

inline bool _trace(_In_z_ _Printf_format_string_ wchar_t const* format, ...)
{
    std::wstring buffer;
    va_list args;
    va_start(args, format);
    Format(buffer, format, args);
    va_end(args);
    OutputDebugStringW(buffer.c_str());
    return true;
}
