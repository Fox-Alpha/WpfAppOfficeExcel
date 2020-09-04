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
            Map(m => m.Typ);
            Map(m => m.Mandant);//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.FormArt);
            Map(m => m.FormIntern);//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.AufNr);
            Map(m => m.IntPos);//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.UntPos);//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.LagerKey);
            Map(m => m.AnLager);
            Map(m => m.ArtikelNr);

            Map(m => m.SerienNr);
            Map(m => m.Kategorie);
            Map(m => m.PosKat);
            Map(m => m.BelegDatum);
            Map(m => m.BelegZeit);
            Map(m => m.Jahr);
            Map(m => m.Periode);
            Map(m => m.BuchungText);
            Map(m => m.Bemerkung);
            Map(m => m.Benutzer);

            Map(m => m.Menge); 
            Map(m => m.Kontonummer);
            Map(m => m.Kasse);
            Map(m => m.Bon);//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.BonPosition);//.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.EingabeArtikelNr);
            Map(m => m.EingabeMenge); //.TypeConverter<CSVImportCustomInt32Converter>();
            Map(m => m.Einheitspreis);//.TypeConverter<CSVImportCustomFloatConverter>();
            Map(m => m.RSGrund).Name("RS-Grund", "RSGrund");
            Map(m => m.Lieferant);

            Map(m => m.LieferDatum);
            Map(m => m.LieferReferenz);
            Map(m => m.Buchung); //.TypeConverter<CSVImportCustomDateTimeConverter>(); //.TypeConverterOption.NullValues("NULL", "NIL", "?", "").TypeConverterOption.DateTimeStyles( DateTimeStyles.None).TypeConverterOption.CultureInfo(CultureInfo.GetCultureInfo("de-DE"));
            Map(m => m.KontrolliertAm); //.TypeConverter<CSVImportCustomDateTimeConverter>();//.TypeConverterOption.NullValues("NULL", "NIL", "?", "").TypeConverterOption.DateTimeStyles(DateTimeStyles.None).TypeConverterOption.CultureInfo(CultureInfo.GetCultureInfo("de-DE"));
            Map(m => m.KontrolliertDurch);
        }
    }
}
