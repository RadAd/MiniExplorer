#pragma once

#include <string>
#include <cstdarg>

inline void Format(std::string& buffer, _In_z_ _Printf_format_string_ char const* format, va_list args)
{
    int const _Result1 = _vscprintf_l(format, NULL, args);
    //if (buffer.size() < _Result1 + 1)
    buffer.resize(_Result1);
    int const _Result2 = _vsprintf_s_l(const_cast<char*>(buffer.data()), static_cast<size_t>(_Result1) + 1, format, NULL, args);
    _ASSERT(-1 != _Result2);
    _ASSERT(_Result1 == _Result2);
}

inline void Format(std::wstring& buffer, _In_z_ _Printf_format_string_ wchar_t const* format, va_list args)
{
    int const _Result1 = _vscwprintf_l(format, NULL, args);
    //if (buffer.size() < _Result1 + 1)
    buffer.resize(_Result1);
    int const _Result2 = _vswprintf_s_l(const_cast<wchar_t*>(buffer.data()), static_cast<size_t>(_Result1) + 1, format, NULL, args);
    _ASSERT(-1 != _Result2);
    _ASSERT(_Result1 == _Result2);
}

inline void Format(std::string& buffer, _In_z_ _Printf_format_string_ char const* format, ...)
{
    va_list args;
    va_start(args, format);
    Format(buffer, format, args);
    va_end(args);
}

inline void Format(std::wstring& buffer, _In_z_ _Printf_format_string_ wchar_t const* format, ...)
{
    va_list args;
    va_start(args, format);
    Format(buffer, format, args);
    va_end(args);
}
