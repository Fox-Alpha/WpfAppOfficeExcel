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
        public uint Mandant { get; set; }
        public string FormArt { get; set; }
        public int FormIntern { get; set; }
        public string AufNr { get; set; }
        public int IntPos { get; set; }
        public int UntPos { get; set; }
        public string LagerKey { get; set; }
        public string AnLager { get; set; }
        public string ArtikelNr { get; set; }
        public string SerienNr { get; set; }
        public string Kategorie { get; set; }
        public string PosKat { get; set; }
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
        public uint Bon { get; set; }
        public UInt32 BonPosition { get; set; }
        public string EingabeArtikelNr { get; set; }
        public uint EingabeMenge { get; set; }
        public float Einheitspreis { get; set; }
        public string RSGrund { get; set; } //RS-Grund
        public string Lieferant { get; set; }
        public string LieferDatum { get; set; }
        public string LieferReferenz { get; set; }
        public DateTime Buchung { get; set; }
        public string KontrolliertAm { get; set; }
        public string KontrolliertDurch { get; set; }

        public CSVImportModel()
        {
        }
    }
}
