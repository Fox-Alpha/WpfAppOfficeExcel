using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WpfAppOfficeExcel.Importer;
using WpfAppOfficeExcel.Models;

namespace WpfAppOfficeExcel
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private ImportOptions import;
        public ImportOptions Import 
        { 
            get => import; 
            set => import = value; 
        }
        public bool BEnableImportOptions 
        { 
            get => bEnableImportOptions; 
            set 
            { 
                bEnableImportOptions = value;
                OnPropertyRaised("BEnableImportOptions");
            } 
        }

        //private CsvDataReader csvDataReader;
        private CsvReader csvFileReader;

        private CSVImportInfoModel _importInfo;

        public CSVImportInfoModel ImportInfo
        {
            get { return _importInfo; }
            set { _importInfo = value; }
        }

        private bool bEnableImportOptions = false;

        private BackgroundWorker worker = new BackgroundWorker();

        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindow()
        {
            InitializeComponent();

            //BackgroundWorker Task
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;

            Import = new ImportOptions();

            this.DataContext = this;

            ImportInfo = new CSVImportInfoModel("", "Export.xlsx");
        }

        public void OnPropertyRaised(string propName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }

        

        private void ButtonFileOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = ".CSV",
                Title = "Eine Datei für den Import auswählen",
                CheckFileExists = true,
                Filter = "Csv files (*.csv)|*.csv|Alle Dateien (*.*)|*.*",
                FilterIndex = 0
            };

            if (openFileDialog.ShowDialog() == true)
            {
                //gbSelectOptionForImport.IsEnabled = true;
                BEnableImportOptions = true;
                //tbFilePathInfo.Text = openFileDialog.FileName;
                ImportInfo.ImportFileName = openFileDialog.FileName;
            }
        }

        private void ButtStartImport_Click(object sender, RoutedEventArgs e)
        {
            pbStatus.Value = 0;
            
            if (worker != null && !worker.IsBusy)
            {
                worker.RunWorkerAsync();
            }
            else
                BEnableImportOptions = false;
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            CsvConfiguration csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture) { AllowComments = true, Delimiter = ";", HasHeaderRecord = true, TrimOptions = TrimOptions.InsideQuotes | TrimOptions.Trim, Encoding = Encoding.Default };

            //using (csvDataReader = new CsvDataReader(new CsvReader(new StreamReader(ImportFileName), csvConfig)))
            //{
            //    csvDataReader.FieldCount;
            //    csvDataReader.get
            //}

            using (csvFileReader = new CsvReader(new StreamReader(ImportInfo.ImportFileName), csvConfig))
            {
                (sender as BackgroundWorker).ReportProgress(0, "Daten Import");
                csvFileReader.Configuration.RegisterClassMap<CSVImportMap>();

                csvFileReader.Read();
                csvFileReader.ReadHeader();

                var recList = csvFileReader.GetRecords<CSVImportModel>().ToList();

                (sender as BackgroundWorker).ReportProgress(15, "Daten Extrahieren");

                //Extrahieren der Filialen
                (sender as BackgroundWorker).ReportProgress(20, "Filialen Extrahieren");
                var Filialen = recList.Select(l => l.LagerKey).GroupBy(x => x)
                             .Where(g => g.Count() > 1)
                             .Select(g => g.Key)
                             .ToList();
                Filialen.Sort();

                List<List<CSVImportModel>> FilialenExport = new List<List<CSVImportModel>>();

                
                foreach (var filiale in Filialen)
                {
                    //ToDo: Filter auf Formular Auswahl setzen
                    var FilOut1 = recList.Select(l => l).Where(w => w.LagerKey == filiale && w.FormArt == "WA").ToList();

                    if (FilOut1.Count > 0)
                    {
                        FilialenExport.Add(FilOut1);
                    }
                }

                /*
                 * Excel Export mit ClosedXML
                 * Datei muss existieren
                 */

                (sender as BackgroundWorker).ReportProgress(60, "Export zu Excel");
                using (var workbook = new XLWorkbook())
                {
                    foreach (var item in Filialen)
                    {
                        var worksheet = workbook.Worksheets.Add(item);
                        int index = Filialen.IndexOf(item);

                        worksheet.Cell(1, 1).InsertData(csvFileReader.Context.HeaderRecord);
                        worksheet.Cell(2, 1).InsertData(FilialenExport[index]);
                    }

                    (sender as BackgroundWorker).ReportProgress(90, "Speichern der Exportdatei");
                    workbook.SaveAs(ImportInfo.ExportFileName);
                }
                /*
                 * *****************************************************************
                 */

                (sender as BackgroundWorker).ReportProgress(95, "Export abgeschlossen");
            }
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbStatus.Value = e.ProgressPercentage;
            pbStatusText.Text = e.UserState as string;
        }

        private void ButtonDebugFile_Click(object sender, RoutedEventArgs e)
        {
            BEnableImportOptions = true;
            
            ImportInfo.ImportFileName = @"C:\Temp\Wolsdorff\Excel_Export_Macro\Tagesbericht-WT5_1010-1015_20200801-20200810.csv";
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pbStatus.Value = 100;
            ButtonOpenExcelExport.IsEnabled = true;
            BEnableImportOptions = true;
        }

        private void ButtonOpenExcelExport_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(ImportInfo.ExportFileName))
            {
                Process pExcelExport = new Process() 
                { 
                    StartInfo = new ProcessStartInfo() 
                    { 
                        FileName = ImportInfo.ExportFileName, 
                        UseShellExecute = true, 
                        Verb = "Open"
                    } 
                };

                pExcelExport.Start();
            }
        }
    }
}
