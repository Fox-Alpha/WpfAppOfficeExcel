using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
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
            get => bEnableImportOptions; set 
            { 
                bEnableImportOptions = value;
                OnPropertyRaised("BEnableImportOptions");
            } 
        }

        private CsvDataReader csvDataReader;
        private CsvReader csvFileReader;


        private string importFileName;
        public string ImportFileName 
        {
            get { return importFileName; }
            private set
            {
                importFileName = value;
            } 
        }

        string ExportToXLSFile = "Export.xlsx";

        public void OnPropertyRaised(string propName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }

        private bool bEnableImportOptions = false;

        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindow()
        {
            InitializeComponent();

            Import = new ImportOptions();
            
            this.DataContext = this;
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
                tbFilePathInfo.Text = openFileDialog.FileName;
                ImportFileName = openFileDialog.FileName;
            }
        }

        private void ButtStartImport_Click(object sender, RoutedEventArgs e)
        {
            CsvConfiguration csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture) { AllowComments = true, Delimiter = ";", HasHeaderRecord = true, TrimOptions = TrimOptions.InsideQuotes | TrimOptions.Trim, Encoding = Encoding.Default };

            //using (csvDataReader = new CsvDataReader(new CsvReader(new StreamReader(ImportFileName), csvConfig)))
            //{
            //    csvDataReader.FieldCount;
            //    csvDataReader.get
            //}

            using (csvFileReader = new CsvReader(new StreamReader(ImportFileName), csvConfig))
            {
                int i = 0;
                csvFileReader.Configuration.RegisterClassMap<CSVImportMap>();

                csvFileReader.Read();
                csvFileReader.ReadHeader();

                var recList = csvFileReader.GetRecords<CSVImportModel>().ToList();

                var Filialen = recList.Select(l => l.LagerKey).GroupBy(x => x)
                             .Where(g => g.Count() > 1)
                             .Select(g => g.Key)
                             .ToList();
                Filialen.Sort();

                List<List<CSVImportModel>> FilialenExport = new List<List<CSVImportModel>>();

                foreach (var filiale in Filialen)
                {
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

                using (var workbook = new XLWorkbook(ExportToXLSFile, ))
                {
                    foreach (var item in Filialen)
                    {
                        var worksheet = workbook.Worksheets.Add(item);
                        int index = Filialen.IndexOf(item);

                        worksheet.Cell(1, 1).InsertData(FilialenExport[index]);
                            //.Cell("A1").Value = "Hello World!";
                        
                    }
                    workbook.Save();
                }


                /*
                 * Excel Export mit CSVHelper Erweiterung
                 * Kann nicht in verschiedenen Sheets der gleichen Datei schreiben
                 */
                //using (var xlsSerializer = new ExcelSerializer(ExportToXLSFile, csvConfig))
                //{
                //    var writer = new CsvWriter(xlsSerializer);
                //    foreach (var item in Filialen)
                //    {
                //        xlsSerializer.Workbook.AddWorksheet(item);
                //        xlsSerializer.Workbook.Worksheets.Worksheet(item);

                //        //{
                //            writer.Configuration.AutoMap<CSVImportModel>();
                //            int index = Filialen.IndexOf(item);
                //            writer.WriteRecords(FilialenExport[index]);
                //        //}
                //    }
                //}
            }
        }
    }
}
