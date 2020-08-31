using System;
using System.Windows.Data;

namespace WpfAppOfficeExcel.Converter
{
    [ValueConversion(typeof(bool), typeof(ImportOptions.enumImportOptions))]
    public class BooleanToImportOptionsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            switch (value)
            {
                case true:
                    return true;
                case false:
                    return false;
                default:
                    return ImportOptions.enumImportOptions.None;
            }
            return ImportOptions.enumImportOptions.None;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is bool)
            {
                if ((bool)value == true)
                    return "yes";
                else
                    return "no";
            }
            return "no";
        }
    }
}
