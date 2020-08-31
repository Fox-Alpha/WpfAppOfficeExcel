using System;
using System.Windows.Data;
using WpfAppOfficeExcel.Importer;

namespace WpfAppOfficeExcel.Converter
{
    //[ValueConversion(typeof(bool), typeof(ImportOptions.enumImportOptions))]
    public class BooleanToImportOptionsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is enumImportOptions)
            {
                if (((enumImportOptions)value & (enumImportOptions)parameter) == (enumImportOptions)parameter)
                    return true;

                if ((enumImportOptions)value == enumImportOptions.None)
                {
                    return false;
                }
                    
            }
            enumImportOptions imp;
            var t = Enum.TryParse(parameter.ToString(), true, out imp);

            if (!t)
            {
                return enumImportOptions.None;
            }

            switch (imp)
            {
                case enumImportOptions.None:
                    return false;
                case enumImportOptions.WarenEingang:
                case enumImportOptions.WarenAusgang:
                case enumImportOptions.ProduktVerlauf:
                case enumImportOptions.ProduktRetoure:
                case enumImportOptions.WarenbewegungPositiv:
                case enumImportOptions.WarenbewegungNegativ:
                case enumImportOptions.UmlagerungEingang:
                case enumImportOptions.UmlagerungAusgang:
                case enumImportOptions.Inventur:
                case enumImportOptions.RabattKunde:
                    if (targetType == typeof(bool?))
                    {
                        var y = ((enumImportOptions)value & (enumImportOptions)parameter) == (enumImportOptions)parameter;
                            return y;
                    }
                    return true;
                default:
                    return false;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is bool)
            {
                if ((bool)value == true)
                {
                    enumImportOptions imp;
                    var t = Enum.TryParse(parameter.ToString(), true, out imp);

                    return t ? imp : enumImportOptions.None;
                }
                else if((bool)value == false)
                {
                    return targetType == typeof(enumImportOptions) ? parameter : enumImportOptions.None;
                }
                return false;
            }
            return enumImportOptions.None;
        }
    }
}
