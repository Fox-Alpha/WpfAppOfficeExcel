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
using System.Windows.Threading;
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
        public DateTime dtTimer { get; set; }
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

        private bool openExportAfterSave;

        public bool OpenExportAfterSave
        {
            get { return openExportAfterSave; }
            set { openExportAfterSave = value; }
        }


        private CSVImportInfoModel _importInfo;

        public CSVImportInfoModel ImportInfo
        {
            get { return _importInfo; }
            set { _importInfo = value; }
        }

        private bool bEnableImportOptions = false;

        private BackgroundWorker worker = new BackgroundWorker();
        private List<string[]> ErrStrLst = new List<string[]>();

        public event PropertyChangedEventHandler PropertyChanged;

        public DispatcherTimer dt;

        public MainWindow()
        {
            InitializeComponent();
            dt = new DispatcherTimer();
            dt.Interval = new TimeSpan(0, 0, 1);
            dt.Tick += TimerTick;

            //BackgroundWorker Task
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;

            Import = new ImportOptions();

            this.DataContext = this;
            dtTimer = new DateTime(1977, 12, 2, 0, 0, 0);

            ImportInfo = new CSVImportInfoModel("", "Export.xlsx");
        }

        private void TimerTick(object sender, EventArgs e)
        {
            dtTimer = dtTimer.AddSeconds(1);
            pbStatusTextTime.Text = dtTimer.ToString("HH:mm:ss");
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
            ButtonOpenExcelExport.IsEnabled = false;

            dtTimer = new DateTime(1977, 12, 2, 0, 0, 0, DateTimeKind.Local);

            if (worker != null && !worker.IsBusy)
            {
                if (dt != null && !dt.IsEnabled)
                {
                    dt.Start();
                }

                pbStatusRun.IsIndeterminate = true;
                worker.RunWorkerAsync();
            }
            else
                BEnableImportOptions = false;
        }

        private bool ReadExceptionResponse(CsvHelperException re)
        {
            //var t = re.InnerException.Data["CsvHelper"].ToString();
            var dat = re.Data;
            var msg = re.Message;
            //var idx = re.ReadingContext.ReusableMemberMapData.Index;
            //var map = re.ReadingContext.ReusableMemberMapData.Names[0];
            var row = re.ReadingContext.Row;
            var rec = re.ReadingContext.Record;
            //var bld = re.ReadingContext.RawRecordBuilder.ToString();

            ErrStrLst.Add(rec);

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Beim Lesen der Daten aus der Import Datei ist ein Fehler aufgetreten");
            sb.AppendLine($"Original Fehler: {msg}");
            sb.AppendLine("");
            sb.AppendLine($"Der Fehler trat in der Zeile {row} auf");
            //sb.AppendLine($"Der Wert: {txt} konnte nicht in das erwartete Format umgewandelt werden");
            //sb.AppendLine($"Der Wert trat in der Spalte {idx + 1} auf");
            //sb.AppendLine($"Der Name der Spalte lautet {map}");
            sb.AppendLine($"");
            sb.AppendLine($"Die Zeile besteht aus folgenden Daten:");
            if (rec.Length > 0)
            {
                sb.AppendLine($"\t{string.Join(", ", rec)}");
            }
            sb.AppendLine($"");
            sb.AppendLine($"");
            sb.AppendLine($"");
            sb.AppendLine($"");
            sb.AppendLine($"");

            var err = sb.ToString();
            sb.Clear();
            sb = null;

            return false;
        }



        private void ButtonDebugFile_Click(object sender, RoutedEventArgs e)
        {
            BEnableImportOptions = true;

            ImportInfo.ImportFileName = @"C:\Temp\Wolsdorff\Excel_Export_Macro\Tagesbericht-WT5_1010-1015_20200801-20200810.csv";
        }

        private void ButtonOpenExcelExport_Click(object sender, RoutedEventArgs e)
        {
            OpenExcelExport();
        }

        private void OpenExcelExport()
        {
            if (File.Exists(ImportInfo.ExportFileName))
            {
                Process pExcelExport = new Process()
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = ImportInfo.ExportFileName,
                        UseShellExecute = true,
                        WindowStyle = ProcessWindowStyle.Maximized,
                        Verb = "Open"
                    }
                };

                pExcelExport.Start();
            }
        }

        private void ButtStartImport_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var bin = (sender as Button).GetBindingExpression(e.Property);
            var obj = bin.DataItem;
        }
    }
}