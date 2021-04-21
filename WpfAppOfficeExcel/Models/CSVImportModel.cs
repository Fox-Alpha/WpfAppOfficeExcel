using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Navigation;

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
        public float Menge { get; set; }
        public string Kontonummer { get; set; }
        public string Kasse { get; set; }
        public string Bon { get; set; }
        public string BonPosition { get; set; }
        public string EingabeArtikelNr { get; set; }
        private int? eingabeMenge;
        public string EingabeMenge
        {
            get { return eingabeMenge.Value.ToString(); }
            set
            {
                if (value != "?" && !string.IsNullOrEmpty(value))
                {
                    if (int.TryParse(value, out int em))
                        eingabeMenge = em;
                    else
                        eingabeMenge = null;
                }
                else
                    eingabeMenge = null;
            }
        }
        public float Einheitspreis { get; set; }
        public string RSGrund { get; set; } //RS-Grund
        public string Lieferant { get; set; }
        public string LieferDatum { get; set; }
        public string LieferReferenz { get; set; }
        public string Buchung { get; set; }

        private static readonly CultureInfo deDE = new CultureInfo("de-DE");
        private DateTime kontrolliertAm;
        public string KontrolliertAm
        //{ get; set; }
        {
            get { return kontrolliertAm.ToString("dd.MM.yyyy"); }
            set
            {
                if (value != "?" && !string.IsNullOrEmpty(value))
                {
                    kontrolliertAm = DateTime.ParseExact(value, @"yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", deDE);
                    //yyyy'-'MM'-'dd'T'HH':'mm':'ss
                }
                else
                    kontrolliertAm = new DateTime(1977, 12, 2);


            }
        }
        public string KontrolliertDurch { get; set; }

        public CSVImportModel()
        {
        }
    }
}
