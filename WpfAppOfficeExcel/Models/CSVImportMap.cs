using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfAppOfficeExcel.Models.Converter;


namespace WpfAppOfficeExcel.Models
{
    public class CSVImportMap : ClassMap<CSVImportModel>
    {
        public CSVImportMap()
        {
            Map(m => m.Typ).Ignore();
            Map(m => m.Mandant).Ignore();//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.FormArt);
            Map(m => m.FormIntern).Ignore();//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.AufNr);
            Map(m => m.IntPos).Ignore();//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.UntPos).Ignore();//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.LagerKey);
            Map(m => m.AnLager).Ignore();
            Map(m => m.ArtikelNr);

            Map(m => m.SerienNr).Ignore();
            Map(m => m.Kategorie).Ignore();
            Map(m => m.PosKat).Ignore();
            Map(m => m.BelegDatum); //.Ignore();
            Map(m => m.BelegZeit).Ignore();
            Map(m => m.Jahr).Ignore();
            Map(m => m.Periode).Ignore();
            Map(m => m.BuchungText);
            Map(m => m.Bemerkung).Ignore();
            Map(m => m.Benutzer).Ignore();

            Map(m => m.Menge).ConvertUsing(row => row.GetField<float>("Menge")); //.TypeConverter<CSVImportCustomFloatConverter>();//
            Map(m => m.Kontonummer).Ignore();
            Map(m => m.Kasse).Ignore();
            Map(m => m.Bon);//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.BonPosition).Ignore();//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.EingabeArtikelNr);
            Map(m => m.EingabeMenge); //.ConvertUsing(row => row.GetField<string>("EingabeMenge")); //<CSVImportCustomInt32Converter>();
            Map(m => m.Einheitspreis).TypeConverter<CSVImportCustomFloatConverter>(); //.ConvertUsing(row => row.GetField<float>("Einheitspreis"));  //.TypeConverter<CSVImportCustomFloatConverter>();
            Map(m => m.RSGrund).Name("RS-Grund", "RSGrund").Ignore();
            Map(m => m.Lieferant);

            Map(m => m.LieferDatum).Ignore();
            Map(m => m.LieferReferenz).Ignore();
            Map(m => m.Buchung); //.TypeConverter(DateTimeConverter);
            //.TypeConverter<CSVImportCustomDateTimeConverter>(); //.TypeConverterOption.NullValues("NULL", "NIL", "?", "").TypeConverterOption.DateTimeStyles( DateTimeStyles.None).TypeConverterOption.CultureInfo(CultureInfo.GetCultureInfo("de-DE"));
            Map(m => m.KontrolliertAm);
            Map(m => m.KontrolliertDurch);//.Ignore();
        }
    }
}
