using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2016.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using WpfAppOfficeExcel.Models;

namespace WpfAppOfficeExcel
{
    public partial class MainWindow
    {
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Encoding enc = Encoding.GetEncoding(1252);
            CsvConfiguration csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                AllowComments = true,
                Delimiter = ";",
                HasHeaderRecord = true,
                TrimOptions = TrimOptions.InsideQuotes | TrimOptions.Trim,
                Encoding = enc,
                BadDataFound = BadDataResponse, 
                ReadingExceptionOccurred = ReadExceptionResponse 
            };

            if (ErrStrLst == null)
            {
                ErrStrLst = new List<string[]>();
            }

            List<CSVImportModel> recList = null;
            List<string> HeaderList = new List<string>();

            using (CsvReader csvFileReader = new CsvReader(new StreamReader(ImportInfo.ImportFileName), csvConfig))
            {
                (sender as BackgroundWorker).ReportProgress(0, "Start Daten Import");

                csvFileReader.Configuration.RegisterClassMap<CSVImportMap>();                

                try
                {
                    csvFileReader.Read();
                    csvFileReader.ReadHeader();
                    HeaderList = (csvFileReader.Context.HeaderRecord.ToList());
                    recList = csvFileReader.GetRecords<CSVImportModel>().ToList();
                }
                catch (CsvHelper.CsvHelperException ex)
                {
                    throw new CsvHelperException(ex.ReadingContext);
                }
            }
            (sender as BackgroundWorker).ReportProgress(15, "Daten Extrahieren");

            if (recList == null || recList.Count == 0)
            {
                (sender as BackgroundWorker).ReportProgress(100, "Fehler beim Daten Extrahieren");
                if (recList != null)
                {
                    recList.Clear();
                    recList = null;
                }
                if (HeaderList != null)
                {
                    HeaderList.Clear();
                    HeaderList = null;
                }
                return;
            }

            /*
            * Extrahieren der Filialen ohne doppelre Einträge
            */

            (sender as BackgroundWorker).ReportProgress(20, "Filialen Extrahieren und sortieren");

            var Filialen = recList.Select(l => l.LagerKey).GroupBy(x => x)
                            .Where(g => g.Count() > 1)
                            .Select(g => g.Key)
                            .ToList();

            Filialen.Sort();

            /************************************************/

            (sender as BackgroundWorker).ReportProgress(30, "Filtern der Daten nach Auswahl");

            List<List<CSVImportModel>> FilialenExport = new List<List<CSVImportModel>>();

            List<string> ImportOptionShortList;
            if ((ImportOptionShortList = Import.GetImportOptionsAsList()).Count == 0)
            {
                (sender as BackgroundWorker).ReportProgress(100, "Fehler: Keine Importoptionen ausgwählt");
                return;
            }

            /*
            * Daten für jede Filiale mit jedem Filterschlüssel extrahieren
            */
            foreach (var filiale in Filialen)
            {
                List<CSVImportModel> FilialExportDaten = new List<CSVImportModel>();

                foreach (var ImportOptionName in ImportOptionShortList)
                {
                    var FilOut = recList.Select(l => l).Where(w => w.LagerKey == filiale && w.FormArt == ImportOptionName).ToList();
                    FilialExportDaten.AddRange(FilOut);
                }

                if (FilialExportDaten.Count > 0)
                {
                    FilialenExport.Add(FilialExportDaten);
                }
                else
                    FilialenExport.Add(new List<CSVImportModel>() { new CSVImportModel() { LagerKey = filiale, Bemerkung = "Keine Daten vorhanden" } });
            }
                
            /*
            * **********************************************************************
            */

            /*
            * Excel Export mit ClosedXML
            * Datei muss existieren
            */

            (sender as BackgroundWorker).ReportProgress(60, "Export zu Excel");

            using (var workbook = new XLWorkbook())
            {
                Debug.WriteLine($"Filialen: {Filialen.Count} == Exports: {FilialenExport.Count}");
                foreach (var item in Filialen)
                {
                    var worksheet = workbook.Worksheets.Add(item);
                    int index = Filialen.IndexOf(item);

                    var rowHeader = worksheet.FirstRow();
                    worksheet.Cell(1, 1).InsertData(HeaderList, true);//csvFileReader.Context.HeaderRecord.ToList(), true);
                    worksheet.Cell(2, 1).InsertData(FilialenExport[index]);

                    if (DeleteUnusedColoums(worksheet))
                    {
                        worksheet.Column( 1).Cell(1).Value = "Buchungstyp";
                        worksheet.Column( 3).Cell(1).Value = "Filiale";
                        worksheet.Column( 5).Cell(1).Value = "Bezeichnung";
                        worksheet.Column( 6).Cell(1).Value = "Summe";
                        worksheet.Column( 8).Cell(1).Value = "Eingabe Artikel Nr. EAN";
                        worksheet.Column(10).Cell(1).Value = "Einzelpreis";
                        worksheet.Column(12).Cell(1).Value = "Buchungs Datum";

                        worksheet.SetAutoFilter();
                        var cc = worksheet.LastCellUsed(); //.CellCount();
                        worksheet.Columns().AdjustToContents(1, cc.Address.RowNumber);

                        //worksheet.Columns("A:XFD").Select();
                        //var x = worksheet.SelectedRanges;
                        
                        worksheet.Sort("D, B, C, E");

                        
                            //(worksheet.Columns().Count());




                        //Formatieren
                        //Druckbereich
                        //Druckeigenschaften
                        //Sortieren

                    }
                }

                (sender as BackgroundWorker).ReportProgress(80, "Speichern Fehlerhafter Zeilen");

                if (ErrStrLst.Count > 0)
                {
                    int row = 2;
                    var worksheet = workbook.Worksheets.Add("Fehlerliste");
                    worksheet.Cell(1, 1).InsertData(HeaderList, true); // csvFileReader.Context.HeaderRecord.ToList(), true);

                    foreach (var item in ErrStrLst)
                    {
                        worksheet.Cell(row, 1).InsertData(item.ToList(), true);
                        row++;
                    }
                }

                (sender as BackgroundWorker).ReportProgress(90, "Speichern der Exportdatei");

                /*
                * Aufräumen der Objecte und freigeben von Speicher
                * ToDo: Encoding bei speichern korrigieren
                */
                using (StreamWriter sw = new StreamWriter(ImportInfo.ExportFileName, false, enc))
                {
                    
                    workbook.SaveAs(sw.BaseStream, new SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false });
                }

                ErrStrLst.Clear();
                HeaderList.Clear();
                recList.Clear();
                Filialen.Clear();
                FilialenExport.Clear();
                ImportOptionShortList.Clear();

                ErrStrLst = null;
                HeaderList = null;
                recList = null;
                Filialen = null;
                FilialenExport = null;
                ImportOptionShortList = null;
            }
            /*
            * *****************************************************************
            */

            (sender as BackgroundWorker).ReportProgress(100, "Export abgeschlossen");
        }

        private bool DeleteUnusedColoums(IXLWorksheet ws)
        {
            if (ws != null)
            {
                ws.Column(35).Delete();
                ws.Column(34).Delete();
                ws.Column(33).Delete();

                ws.Column(32).Delete();
                ws.Column(29).Delete();
                ws.Column(25).Delete();
                ws.Column(23).Delete();
                ws.Column(22).Delete();
                ws.Column(20).Delete();

                ws.Column(19).Delete();
                ws.Column(17).Delete();
                ws.Column(16).Delete();
                ws.Column(15).Delete();
                ws.Column(14).Delete();
                ws.Column(13).Delete();
                ws.Column(12).Delete();
                ws.Column(11).Delete();

                ws.Column(9).Delete();
                ws.Column(7).Delete();
                ws.Column(6).Delete();
                ws.Column(4).Delete();
                ws.Column(2).Delete();
                ws.Column(1).Delete();

                

                return true;
            }
            return false;
        }

        private void BadDataResponse(ReadingContext obj)
        {
            int row = obj.Row;
            string col = obj.Field;
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbStatus.Value = e.ProgressPercentage;
            pbStatusText.Text = e.UserState as string;
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)

        {
            pbStatus.Value = 100;
            pbStatusRun.IsIndeterminate = false;
            ButtonOpenExcelExport.IsEnabled = true;
            BEnableImportOptions = true;

            //Timer für Messung stoppen
            if(dt.IsEnabled)
                dt.Stop();
        }
    }
}
