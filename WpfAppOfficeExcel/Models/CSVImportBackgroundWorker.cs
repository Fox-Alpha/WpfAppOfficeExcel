﻿using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfAppOfficeExcel.Models;

namespace WpfAppOfficeExcel
{
    public partial class MainWindow
    {
        //Zähler für den Import/Export Fortschritt
        private int iProgress = 0;

        /// <summary>
        /// Backgroundworker Main Function
        /// </summary>
        /// <param name="sender">Backgroundworker Instanz</param>
        /// <param name="e">Backgroundworker Argumente</param>
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bgworker = sender as BackgroundWorker;

            Encoding enc = Encoding.GetEncoding("iso-8859-1");

            int[] CoulumnsToDelete;
            if (Import.ExpKmpgColumns == true)
            {
                CoulumnsToDelete = new int[]
                {
                    33, 32, 31, 29, 25, 23, 22, 20, 19, 17, 16, 15, 13, 12, 11, 9, 7, 6, 5, 4, 2, 1
                };
            }
            else
            {
                CoulumnsToDelete = new int[]
                {
                    35, 34, 33, 32, 31, 29, 25, 23, 22, 20, 19, 17, 16, 15, 13, 12, 11, 9, 7, 6, 5, 4, 2, 1
                };
            }
            
            int[] IndexToRename = new int[] 
            {
                1, 2, 3, 5, 6, 8, 10, 4
            };
            string[] ColumnNames = new string[]
            {
                "Buchungstyp",
                "Filiale",
                "Warengruppe",
                "Bezeichnung",
                "Summe",
                "Eingabe Artikel Nr. EAN",
                "Einzelpreis",
                "Buchungs Datum"
            };

            List<string[]> ErrStrLst = new List<string[]>();

            List<CSVImportModel> recList;
            List<string> HeaderList = new List<string>();

            bgworker.ReportProgress(iProgress = 3, "Starten des Datenimport. Einlesen der CSV Datei ...");

            if (!CSVImportReadFile(out HeaderList, out recList) || bgworker.CancellationPending)
            {
                e.Cancel = true;
            }

            //Fehler, wenn keine Daten gelesen wurden
            if (recList == null || recList.Count == 0)
            {
                bgworker.ReportProgress(100, "Fehler beim Daten Extrahieren");
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
                e.Cancel = true;    //bgw abruch signalisieren
            }

            if (!e.Cancel && !bgworker.CancellationPending)
            {
                bgworker.ReportProgress(iProgress += 7, "Filialen Extrahieren und sortieren");
                List<string> Filialen = GetFilialListe(recList);
                List<CSVImportModel> FilialenExport = new List<CSVImportModel>();

                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("de-DE");

                using (var workbook = new XLWorkbook(new LoadOptions() { EventTracking = XLEventTracking.Enabled }))
                {
                    int fi = 1;
                    IXLWorksheet worksheet = null;

                    

                    //foreach (var filiale in Filialen)
                    {
                        {
                            bgworker.ReportProgress(iProgress++, "Export zu Excel");

                            if (!import.OneSheetOnly)
                            {

                                foreach (var filiale in Filialen)
                                {
                                    FilialenExport = GetFilialDataForExport(recList, filiale);

                                    if (FilialenExport != null && FilialenExport.Count > 0)
                                    {
                                        bgworker.ReportProgress(iProgress++, $"Filtern der Daten für Filiale {filiale} - {fi}/{Filialen.Count()}");

                                        worksheet = workbook.Worksheets.Add(filiale);

                                        if (worksheet is null)
                                        {
                                            bgworker.ReportProgress(iProgress++, $"Fehler beim anlegen der Excel Daten : {filiale}");
                                            //e.Cancel = true;
                                            //CleanupLists(ref recList, ref Filialen, ref FilialenExport, ref HeaderList);
                                            //return;
                                            //break;
                                            continue;

                                        }

                                        SaveFilialDataToWorksheet(HeaderList, FilialenExport, worksheet);

                                        bgworker.ReportProgress(iProgress++, $"Anpassen der Exportierten Daten: {filiale} - {fi} / {Filialen.Count()}");

                                        if (DeleteUnusedColoumns(worksheet, CoulumnsToDelete))
                                        {
                                            RenameCoulumns(worksheet, IndexToRename, ColumnNames);

                                            SortAndFormatXlsSheet(worksheet);
                                        }
                                    }
                                    else if (FilialenExport != null && FilialenExport.Count == 0)
                                    {
                                        bgworker.ReportProgress(iProgress++, $"Keine Daten für Filiale {filiale} für Export gefunden");
                                    }
                                }
                                if (workbook != null)
                                    SaveWorkBookToFile(e, bgworker, enc, ErrStrLst, HeaderList, workbook);

                                CleanupLists(ref recList, ref Filialen, ref FilialenExport, ref HeaderList);
                                return;
                            }

                            if (import.OneSheetOnly)
                            {
                                worksheet = workbook.Worksheets.Add("Datenexport");
                            }

                            for (int i = 0; i < Filialen.Count; i++)
                            {
                                FilialenExport.AddRange(GetFilialDataForExport(recList, Filialen[i]));
                            }

                            if(FilialenExport == null || FilialenExport.Count == 0)
                            {
                                bgworker.ReportProgress(iProgress++, $"Filtern der Daten für Filialen");
                                CleanupLists(ref recList, ref Filialen, ref FilialenExport, ref HeaderList);
                                e.Cancel = true;
                                return;
                            }

                            SaveFilialDataToWorksheet(HeaderList, FilialenExport, worksheet);

                            //bgworker.ReportProgress(iProgress++, $"Anpassen der Exportierten Daten: {filiale} - {fi} / {Filialen.Count()}");

                            if (DeleteUnusedColoumns(worksheet, CoulumnsToDelete))
                            {
                                RenameCoulumns(worksheet, IndexToRename, ColumnNames);

                                SortAndFormatXlsSheet(worksheet);
                            }

                            FilialenExport.Clear();
                            //worksheet = null;
                            fi++;
                        }
                        
                        //if (FilialenExport != null && FilialenExport.Count == 0)
                        //{
                        //    bgworker.ReportProgress(iProgress++, $"Keine Daten für Filiale {filiale} für Export gefunden");
                        //}
                        //else
                        //{
                        //    e.Cancel = true;
                        //    //break;
                        //}

                        if (e.Cancel || bgworker.CancellationPending)
                        {
                            e.Cancel = true;
                            SaveWorkBookToFile(e, bgworker, enc, ErrStrLst, HeaderList, workbook);
                            return;
                            //break;
                        }
                    }
                    /*
                    * Aufräumen der Objecte und freigeben von Speicher
                    */
                    CleanupLists(ref recList, ref Filialen, ref FilialenExport, ref HeaderList);

                    if (workbook != null)
                        SaveWorkBookToFile(e, bgworker, enc, ErrStrLst, HeaderList, workbook);
                }
            }
            ErrStrLst.Clear();
            ErrStrLst = null;
        }

        private void CleanupLists(ref List<CSVImportModel> recList, ref List<string> Filialen, ref List<CSVImportModel> FilialenExport, ref List<string> HeaderList)
        {
            recList.Clear();
            recList = null;

            Filialen.Clear();
            FilialenExport?.Clear();
            Filialen = null;
            FilialenExport = null;

            HeaderList.Clear();
            HeaderList = null;
        }

        private void SaveWorkBookToFile(DoWorkEventArgs e, BackgroundWorker bgworker, Encoding enc, List<string[]> ErrStrLst, List<string> HeaderList, XLWorkbook workbook)
        {
            bgworker.ReportProgress(iProgress++, "Speichern Fehlerhafter Zeilen");
            SaveErrorData(ErrStrLst, HeaderList, workbook);

            if (!e.Cancel && !bgworker.CancellationPending)
            {
                bgworker.ReportProgress(iProgress++, "Speichern der Exportdatei");

                using (StreamWriter sw = new StreamWriter(ImportInfo.ExportFileName, false, enc))
                {
                    workbook.SaveAs(sw.BaseStream,
                                    new SaveOptions
                                    {
                                        EvaluateFormulasBeforeSaving = false,
                                        GenerateCalculationChain = false,
                                        ValidatePackage = false
                                    });
                    bgworker.ReportProgress(iProgress = 100, "Export abgeschlossen");
                }
            }
        }

        /// <summary>
        /// Einlesen der Rohdaten in den Speicher zum weiteren bearbeiten
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="HeaderList"></param>
        /// <param name="recList"></param>
        /// <returns></returns>
        private bool CSVImportReadFile(out List<string> HeaderList, out List<CSVImportModel> recList)
        {
            bool hasData = false;
            recList = new List<CSVImportModel>();
            HeaderList = new List<string>();
            Encoding enc = Encoding.GetEncoding("iso-8859-1");

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

            using (CsvReader csvFileReader = new CsvReader(new StreamReader(ImportInfo.ImportFileName), csvConfig))
            {
                csvFileReader.Configuration.RegisterClassMap<CSVImportMap>();

                try
                {
                    csvFileReader.Read();
                    csvFileReader.ReadHeader();
                    HeaderList = csvFileReader.Context.HeaderRecord.ToList();
                    recList = csvFileReader.GetRecords<CSVImportModel>().ToList();
                    hasData = true;
                }
                catch (CsvHelper.CsvHelperException ex)
                {
                    throw new CsvHelperException(ex.ReadingContext);
                }
            }
            return hasData;
        }

        /// <summary>
        /// Die Filialliste aus den Daten erstellen
        /// Zum anlegen der einzelnen Arbeitsblätter
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="recList"></param>
        /// <returns></returns>
        private List<string> GetFilialListe(List<CSVImportModel> recList)
        {
            var Filialen = recList.Select(l => l.LagerKey).GroupBy(x => x)
                            .Where(g => g.Count() > 1)
                            .Select(g => g.Key)
                            .ToList();

            Filialen.Sort();

            ImportInfo.AnzahlFiliale = Filialen.Count();
            return Filialen;
        }

        /// <summary>
        /// Daten für jede Filiale mit jedem Filterschlüssel extrahieren
        /// </summary>
        /// <param name="sender">Der Backgroundworker</param>
        /// <param name="filiale">Die zu bearbeitende Filiale</param>
        /// <returns></returns>
        private List<CSVImportModel> GetFilialDataForExport(List<CSVImportModel> Data, string filiale)
        {
            List<CSVImportModel> FilialExportDaten = new List<CSVImportModel>();

            List<string> ImportOptionShortList;

            //TODO: Abbruchbedingung überarbeiten
            if ((ImportOptionShortList = Import.GetImportOptionsAsList()).Count == 0)
            {
                ImportOptionShortList.Clear();
                ImportOptionShortList = null;
                FilialExportDaten = null;

                //if ((sender as BackgroundWorker).WorkerSupportsCancellation)
                //    (sender as BackgroundWorker).CancelAsync();

                return null; ;
            }

            foreach (var ImportOptionName in ImportOptionShortList)
            {
                if (string.IsNullOrWhiteSpace(ImportOptionName))
                {
                    continue;
                }
                var FilOut = Data.Select(l => l).Where(w => w.LagerKey == filiale && w.FormArt == ImportOptionName).ToList();

                if (FilOut.Count > 0)
                {
                    FilialExportDaten.AddRange(FilOut);

                }
                else
                    FilialExportDaten.Add(new CSVImportModel() { FormArt = ImportOptionName, LagerKey = filiale, BuchungText = "Keine Daten vorhanden" });
            }

            ImportOptionShortList.Clear();
            ImportOptionShortList = null;

            return FilialExportDaten;
        }

        /// <summary>
        /// Speichern der gefilterten Daten einer Filiale in einem eigenen Arbeitsblatt
        /// </summary>
        /// <param name="HeaderList"></param>
        /// <param name="FilialenExport"></param>
        /// <param name="worksheet"></param>
        private void SaveFilialDataToWorksheet(List<string> HeaderList, List<CSVImportModel> FilialenExport, IXLWorksheet worksheet)
        {
            if (HeaderList != null) worksheet.Cell(1, 1).InsertData(HeaderList, true);
            if (FilialenExport != null) worksheet.Cell(2, 1).InsertData(FilialenExport);
        }

        /// <summary>
        /// Nicht mehr benötigte Spalten entfernen
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="IndexToDelete"></param>
        /// <returns></returns>
        private bool DeleteUnusedColoumns(IXLWorksheet ws, int[] IndexToDelete)
        {
            if (ws != null && IndexToDelete.Length > 0)
            {
                foreach (var item in IndexToDelete)
                {
                    ws.Column(item).Delete();
                }
                return true;
            }
            return false;
        }
        
        /// <summary>
         /// Umbenennen der SPalten zur besseren Lesbarkeit
         /// </summary>
         /// <param name="worksheet"></param>
        private bool RenameCoulumns(IXLWorksheet worksheet, int[] IndexToRename, string[] ColumnNames)
        {
            //Dictionary<string, string> debug = new Dictionary<string, string>();
            //Umbenennen der Überschriften
            if (IndexToRename.Length != ColumnNames.Length)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < IndexToRename.Length; i++)
                {
                    //debug.Add(worksheet.Column(IndexToRename[i]).Cell(1).Value.ToString(), ColumnNames[i]);
                    worksheet.Column(IndexToRename[i]).Cell(1).Value = ColumnNames[i];
                }
            }
            //debug.Clear();
            //debug = null;
            return true;
        }

        /// <summary>
        /// Sortieren der Daten
        /// Überschriften auf BOLD setzen
        /// Autofilter Aktivieren
        /// Spaltenbreite an Inhalt anpassen
        /// </summary>
        /// <param name="worksheet"></param>
        private void SortAndFormatXlsSheet(IXLWorksheet worksheet)
        {
            ////Sortieren der Daten
            var lastCellUsed = worksheet.LastCellUsed();
            var lastCellUsedAddress = $"A2:{lastCellUsed.Address}";
            var DataRange = worksheet.Range(lastCellUsedAddress);
            DataRange.Sort("F, C, D, E");

            //Autofilter und Spaltenbreite an Inhalt anpassen
            worksheet.RangeUsed().SetAutoFilter();

            lastCellUsed = worksheet.LastCellUsed();
            //Spaltenbreite an Inhalt anpassen
            worksheet.Columns().AdjustToContents(1, lastCellUsed.Address.RowNumber);
            worksheet.Row(1).Style.Font.SetBold();

            //Spalten als Zahl formartieren
            worksheet.Columns("F,J").AdjustToContents(1, lastCellUsed.Address.RowNumber).Style.NumberFormat.NumberFormatId = 2;

            worksheet.Range($"I2:I{lastCellUsed.Address.RowNumber}").SetDataType(XLDataType.Number);

            //Formatieren
            //Druckbereich
            //Druckeigenschaften
        }

        /// <summary>
        /// Speichern Fehlerhafter Zeilen zur manuellen Korrektur
        /// </summary>
        /// <param name="HeaderList"></param>
        /// <param name="workbook"></param>
        private void SaveErrorData(List<string[]> ErrStrLst, List<string> HeaderList, XLWorkbook workbook)
        {           
            if (ErrStrLst.Count > 0)
            {
                int row = 2;
                var errorworksheet = workbook.Worksheets.Add("Fehlerliste");

                errorworksheet.Cell(1, 1).InsertData(HeaderList, true);

                foreach (var item in ErrStrLst)
                {
                    errorworksheet.Cell(row, 1).InsertData(item.ToList(), true);
                    row++;
                }
            }
        }

        /// <summary>
        /// Fortschritt an GUI weiterleiten
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbStatus.Value = e.ProgressPercentage;
            pbStatusText.Text = e.UserState as string;

            WriteLogToFile($"{e.ProgressPercentage} - {e.UserState as string}");
        }

        private void WriteLogToFile(string Message)
        {
            string MsgLog = string.Format($"{DateTime.Now.ToString("T")}: {Message}\r\n");

            File.AppendAllText("MessageLog.txt", MsgLog);
        }

        /// <summary>
        /// Vorgang ist komplett. Der Worker beendet
        /// Aufräumen
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pbStatusRun.IsIndeterminate = false;
            pbStatus.Value = pbStatus.Maximum;

            if (e.Cancelled)
            {
                pbStatusText.Text = "Abgebrochen";
                WriteLogToFile("Abgebrochen");
                pbStatusRun.IsIndeterminate = false;
            }
            else if (e.Error != null)
            {
                pbStatusText.Text = "Error: " + e.Error.Message;
                WriteLogToFile($"Error: {e.Error.Message}");
                WriteLogToFile($"Error: {e.Error.StackTrace}");
            }
            else
            {
                ButtonOpenExcelExport.IsEnabled = true;
                ButtonSaveExportAs.IsEnabled = true;

                if (OpenExportAfterSave == true)
                {
                    OpenExcelExport();
                }
            }

            ButtCancelImport.IsEnabled = false;
            ButtCancelImport.Visibility = System.Windows.Visibility.Collapsed;

            BEnableImportOptions = true;

            //Timer für Messung stoppen
            if (dt.IsEnabled)
                dt.Stop();
        }

        /// <summary>
        /// Event, wenn Daten nicht richtig gelesen werden konnten
        /// </summary>
        /// <param name="obj"></param>
        private void BadDataResponse(ReadingContext obj)
        {
            int row = obj.Row;
            string col = obj.Field;
            Debug.WriteLine($"{row} -- {col}", "BadDataResponse");
        }
    }
}
