using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WpfAppOfficeExcel.Models;

namespace WpfAppOfficeExcel
{
    public partial class MainWindow
    {
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            CsvConfiguration csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture) { AllowComments = true, Delimiter = ";", HasHeaderRecord = true, TrimOptions = TrimOptions.InsideQuotes | TrimOptions.Trim, Encoding = Encoding.Default, BadDataFound = BadDataResponse, ReadingExceptionOccurred = ReadExceptionResponse };

            //using (csvDataReader = new CsvDataReader(new CsvReader(new StreamReader(ImportFileName), csvConfig)))
            //{
            //    csvDataReader.FieldCount;
            //    csvDataReader.get
            //}

            if (ErrStrLst == null)
            {
                ErrStrLst = new List<string[]>();
            }

            using (CsvReader csvFileReader = new CsvReader(new StreamReader(ImportInfo.ImportFileName), csvConfig))
            {
                (sender as BackgroundWorker).ReportProgress(0, "Start Daten Import");

                csvFileReader.Configuration.RegisterClassMap<CSVImportMap>();
                List<CSVImportModel> recList = null;
                try
                {
                    csvFileReader.Read();
                    csvFileReader.ReadHeader();

                    recList = csvFileReader.GetRecords<CSVImportModel>().ToList();
                }
                catch (CsvHelper.TypeConversion.TypeConverterException re)
                {
                }


                (sender as BackgroundWorker).ReportProgress(15, "Daten Extrahieren");

                if (recList == null)
                {
                    (sender as BackgroundWorker).ReportProgress(100, "Fehler beim Daten Extrahieren");
                    return;
                }

                //Extrahieren der Filialen
                (sender as BackgroundWorker).ReportProgress(20, "Filialen Extrahieren und sortieren");

                var Filialen = recList.Select(l => l.LagerKey).GroupBy(x => x)
                             .Where(g => g.Count() > 1)
                             .Select(g => g.Key)
                             .ToList();

                Filialen.Sort();

                (sender as BackgroundWorker).ReportProgress(30, "Filtern der Daten nach Auswahl");
                List<List<CSVImportModel>> FilialenExport = new List<List<CSVImportModel>>();

                List<string> vs = Import.GetImportOptionsAsList();

                foreach (var filiale in Filialen)
                {
                    //ToDo: Schleife durch ausgewähglte Import Optionen
                    //ToDo: Filter auf Formular Auswahl setzen
                    var FilOut1 = recList.Select(l => l).Where(w => w.LagerKey == filiale && w.FormArt == "WA").ToList();

                    if (FilOut1.Count > 0)
                    {
                        FilialenExport.Add(FilOut1);
                    }
                    else
                        FilialenExport.Add(new List<CSVImportModel>() { new CSVImportModel() { LagerKey = filiale, Bemerkung = "Keine Daten vorhanden" } }) ;
                }

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
                        //rowHeader.Cell(1).InsertData(csvFileReader.Context.HeaderRecord);
                        worksheet.Cell(1, 1).InsertData(csvFileReader.Context.HeaderRecord.ToList(), true);
                        //worksheet.Cell(1, 1).AsRange();
                        ///ToDo: Fehler wenn im Eintrag keine Daten gefüllt sind
                        worksheet.Cell(2, 1).InsertData(FilialenExport[index]);
                    }

                    (sender as BackgroundWorker).ReportProgress(80, "Speichern Fehlerhafter Zeilen");

                    if (ErrStrLst.Count > 0)
                    {
                        int row = 2;
                        var worksheet = workbook.Worksheets.Add("Fehlerliste");
                        worksheet.Cell(1, 1).InsertData(csvFileReader.Context.HeaderRecord.ToList(), true);

                        foreach (var item in ErrStrLst)
                        {
                            worksheet.Cell(row, 1).InsertData(item.ToList(), true);
                            row++;
                        }
                    }

                    (sender as BackgroundWorker).ReportProgress(90, "Speichern der Exportdatei");
                    workbook.SaveAs(ImportInfo.ExportFileName);
                    //workbook.Dispose();
                    ErrStrLst.Clear();
                    recList.Clear();
                    Filialen.Clear();
                    FilialenExport.Clear();
                    
                    //workbook.Dispose();
                    //workbook = null;
                    ErrStrLst = null;
                    recList = null;
                    Filialen = null;
                    FilialenExport = null;
                }
                /*
                 * *****************************************************************
                 */

                (sender as BackgroundWorker).ReportProgress(95, "Export abgeschlossen");
            }
            Debug.WriteLine("The highest generation is {0}", GC.MaxGeneration);
            Debug.WriteLine("Total Memory: {0}", GC.GetTotalMemory(false));
            GC.Collect(0);
            Debug.WriteLine("Total Memory: {0}", GC.GetTotalMemory(false));
            GC.Collect(2);
            Debug.WriteLine("Total Memory: {0}", GC.GetTotalMemory(false));

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
            if(dt.IsEnabled)
                dt.Stop();
        }
    }
}
