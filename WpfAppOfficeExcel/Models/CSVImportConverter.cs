using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using DocumentFormat.OpenXml.Office.CustomUI;

namespace WpfAppOfficeExcel.Models.Converter
{
    public class CSVImportCustomInt32Converter : DefaultTypeConverter
    {
        public override object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            if (text == "?")
            {
                return 0;
            }

            int c; 
            if(int.TryParse(text, out c))
            {
                return c;
            }

            return null; 
                //base.ConvertFromString(text, row, memberMapData);
        }
    }

    public class CSVImportCustomFloatConverter : DefaultTypeConverter
    {
        public override object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            float f;

            if (text == "?")
            {
                return 0.0f;
            }

            if (float.TryParse(text, out f))
            {
                return f;
            }

            return null; 
            //base.ConvertFromString(text, row, memberMapData);    
        }
    }

    public class CSVImportCustomDateTimeConverter : DefaultTypeConverter
    {
        //private string CustomDateFormat = @"yyyy-MM-dd'T'hh:mm:ss'.'FFF";
        //private string CustomDateFormat = @"dd/MM/yyyy";

        public override object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            //DateTime newDate = default;

            if (text == "?" || string.IsNullOrEmpty(text))
            {
                return "n/a";
            }

            try
            {
                //newDate = DateTime.Parse(text, CultureInfo.InvariantCulture);

                //newDate = DateTime.ParseExact(text, CustomDateFormat, CultureInfo.InvariantCulture);

                //GetCultureInfo("de-DE")CurrentCulture.DateTimeFormat
                //newDate = DateTime.ParseExact(text, CustomDateFormat, CultureInfo.InvariantCulture); //); //, CultureInfo.InvariantCulture, DateTimeStyles.None); //, );
                //newDate.
                //2019-04-01T20:08:13.929
                //memberMapData.TypeConverterOptions.CultureInfo = CultureInfo.GetCultureInfo("de-DE");
                //memberMapData.TypeConverterOptions.DateTimeStyle = DateTimeStyles.AssumeLocal;
                //memberMapData.TypeConverterOptions.Formats = new string[] { "dd.MM.yyyy hh:mm:ss"};
                //memberMapData.TypeConverterOptions.


                //DateTime dt = new DateTime(1970,12,01);
                //dt.Kind = DateTimeKind.Local;
                //"01.01.1970 00:00:00"
                //newDate = DateTime.ParseExact("01.01.1970 12:00:00", @"dd.MM.yyyy hh:mm:ss",  CultureInfo.InvariantCulture);
                //  return null; // newDate; // base.ConvertFromString(newDate.ToString(), row, memberMapData); 
            }
            catch (Exception ex)
            {
                Debug.WriteLine(String.Format(@"Error parsing date '{0}': {1}", text, ex.Message));
            }

            return text;

            //return base.ConvertFromString(text, row, memberMapData);    
        }
    }

    public class CSVImportInvalidInt32Converter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }
    }

    public class CSVImportInvalidDateConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }
    }

    public class CSVImportInvalidFloatConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }
    }

    public class CSVImportInvalidDateTimeConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }
    }
}
