using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MSOfficeAIAssistant.UI.Converters
{
    public class BooleanToVisibilityConverter : IValueConverter
    {
        public bool Invert { get; set; }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool boolVal = (value is bool) && (bool)value;
            if (Invert) boolVal = !boolVal;
            return boolVal ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility)
            {
                Visibility v = (Visibility)value;
                bool res = (v == Visibility.Visible);
                return Invert ? !res : res;
            }
            return false;
        }
    }
}
