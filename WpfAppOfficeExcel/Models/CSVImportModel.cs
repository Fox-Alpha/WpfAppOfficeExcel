using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace WpfAppOfficeExcel.Models
{
    public class CSVImportModel
    {
        public string Typ { get; set; }
        public string Mandant { get; set; }
        public string FormArt { get; set; }
        public string FormIntern { get; set; }
        public string AufNr { get; set; }
        public string IntPos { get; set; }
        public string UntPos { get; set; }
        public string LagerKey { get; set; }
        public string AnLager { get; set; }
        public string ArtikelNr { get; set; }
        public string SerienNr { get; set; }
        public string Kategorie { get; set; }
        public string PosKat { get; set; }
        //TODO: Typ Umstellung ggf. Converter erstellen und verwenden
        public string BelegDatum { get; set; }
        public string BelegZeit { get; set; }
        public string Jahr { get; set; }
        public string Periode { get; set; }
        public string BuchungText { get; set; }
        public string Bemerkung { get; set; }
        public string Benutzer { get; set; }
        public string Menge { get; set; }
        public string Kontonummer { get; set; }
        public string Kasse { get; set; }
        public string Bon { get; set; }
        public string BonPosition { get; set; }
        public string EingabeArtikelNr { get; set; }
        public string EingabeMenge { get; set; }
        public string Einheitspreis { get; set; }
        public string RSGrund { get; set; } //RS-Grund
        public string Lieferant { get; set; }
        public string LieferDatum { get; set; }
        public string LieferReferenz { get; set; }
        public string Buchung { get; set; }
        public string KontrolliertAm { get; set; }
        public string KontrolliertDurch { get; set; }

        public CSVImportModel()
        {
        }
    }
}
