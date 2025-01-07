﻿using System.Globalization;
using System.Windows.Data;

namespace XlsxDiffTool.Common;

public class InverseBoolConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return !(value as bool? == true);
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return !(value as bool? == true);
    }

}
