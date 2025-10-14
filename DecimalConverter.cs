using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace GerenciadorViveiro;

public class DecimalConverter : IValueConverter {
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        if (value is decimal decimalValue) {
            return decimalValue.ToString("F2", CultureInfo.CurrentCulture);
        }
        return value?.ToString();
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        if (value is string stringValue) {
            // Substitui ponto por vírgula e tenta fazer o parse
            stringValue = stringValue.Replace(".", ",");
            
            if (decimal.TryParse(stringValue, NumberStyles.Any, CultureInfo.CurrentCulture, out decimal result)) {
                return result;
            }
        }
        return 0m;
    }
}