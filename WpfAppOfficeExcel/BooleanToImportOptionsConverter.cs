using System;
using System.Diagnostics;
using System.Globalization;
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
                if (((enumImportOptions)value & (enumImportOptions)parameter) == (enumImportOptions.None))
                //if ((enumImportOptions)value == enumImportOptions.None)
                {
                    return false;
                }

                if (((enumImportOptions)value & (enumImportOptions)parameter) == (enumImportOptions)parameter)
                    return true;

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
                        var addremove = ((enumImportOptions)value & (enumImportOptions)parameter) == (enumImportOptions)parameter;
                            return addremove;
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
                    var validParam = Enum.TryParse(parameter.ToString(), true, out imp);

                    return validParam ? imp : enumImportOptions.None;
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

    public class IsEnabledImportOptionsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is enumImportOptions)
            {
                var val = enumImportOptions.Inventur | enumImportOptions.ProduktRetoure | enumImportOptions.ProduktVerlauf | enumImportOptions.RabattKunde | enumImportOptions.UmlagerungAusgang | enumImportOptions.UmlagerungEingang | enumImportOptions.WarenAusgang | enumImportOptions.WarenbewegungNegativ | enumImportOptions.WarenbewegungPositiv | enumImportOptions.WarenEingang;

                var isNotNone = EnumsNET.FlagEnums.HasAnyFlags(typeof(enumImportOptions), value, val);
                
                return isNotNone;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return true;
        }
    }
}
