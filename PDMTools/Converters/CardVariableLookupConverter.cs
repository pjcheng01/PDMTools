using System;
using System.Globalization;
using System.Windows.Data;
using PDMTools.Models;

namespace PDMTools.Converters
{
    /// <summary>依 ConverterParameter（變數鍵）從 BomItem.CardVariables 取出顯示文字。</summary>
    public sealed class CardVariableLookupConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var item = value as BomItem;
            var key = parameter as string;
            if (item == null || string.IsNullOrEmpty(key))
            {
                return string.Empty;
            }

            if (item.CardVariables != null && item.CardVariables.TryGetValue(key, out var v))
            {
                return v ?? string.Empty;
            }

            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
