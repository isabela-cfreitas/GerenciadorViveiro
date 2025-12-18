using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace GerenciadorViveiro;

public class IntegerConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is int intValue)
        {
            return intValue.ToString();
        }
        return value?.ToString() ?? string.Empty;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is string stringValue)
        {
            // Remove espaços em branco
            stringValue = stringValue.Trim();
            
            if (string.IsNullOrEmpty(stringValue))
            {
                return 0;
            }
            
            if (int.TryParse(stringValue, NumberStyles.Integer, CultureInfo.CurrentCulture, out int result))
            {
                return result;
            }
            
            // Se não conseguir converter, retorna o valor anterior ou 0
            return 0;
        }
        return 0;
    }
}