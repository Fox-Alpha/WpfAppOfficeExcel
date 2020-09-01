using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfAppOfficeExcel.Models
{
    public class CSVImportMap : ClassMap<CSVImportModel>
    {
        public CSVImportMap()
        {
            Map(m => m.Typ);
            Map(m => m.Mandant);
            Map(m => m.FormArt);
            Map(m => m.FormIntern);
            Map(m => m.AufNr);
            Map(m => m.IntPos);
            Map(m => m.UntPos);
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
            Map(m => m.Bon);
            Map(m => m.BonPosition);
            Map(m => m.EingabeArtikelNr);
            Map(m => m.EingabeMenge);
            Map(m => m.Einheitspreis);
            Map(m => m.RSGrund).Name("RS-Grund", "RSGrund");
            Map(m => m.Lieferant);

            Map(m => m.LieferDatum);
            Map(m => m.LieferReferenz);
            Map(m => m.Buchung);
            Map(m => m.KontrolliertAm);
            Map(m => m.KontrolliertDurch);
        }
    }
}
