using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfAppOfficeExcel.Models
{
    class CSVImportModel
    {
        string Typ { get; set; }
        int Mandant { get; set; }
        string FormArt { get; set; }
        int FormIntern { get; set; }
        string AufNr { get; set; }
        string IntPos { get; set; }
        string UntPos { get; set; }
        string LagerKey { get; set; }
        string AnLager { get; set; }
        string ArtikelNr { get; set; }
        string SerienNr { get; set; }
        string Kategorie { get; set; }
        string PosKat { get; set; }
        string BelegDatum { get; set; }
        string BelegZeit { get; set; }
        string Jahr { get; set; }
        string Periode { get; set; }
        string BuchungText { get; set; }
        string Bemerkung { get; set; }
        string Benutzer { get; set; }
        string Menge { get; set; }
        string Kontonummer { get; set; }
        string Kasse { get; set; }
        string Bon { get; set; }
        string BonPosition { get; set; }
        string EingabeArtikelNr { get; set; }
        string EingabeMenge { get; set; }
        string Einheitspreis { get; set; }
        string RSGrund { get; set; } //RS-Grund
        string Lieferant { get; set; }
        string LieferDatum { get; set; }
        string LieferReferenz { get; set; }
        string Buchung { get; set; }
        string KontrolliertAm { get; set; }
        string KontrolliertDurch { get; set; }

        public CSVImportModel()
        {
        }
    }
}
