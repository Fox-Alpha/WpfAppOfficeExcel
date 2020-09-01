using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
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

                //if (csvFileReader.ReadHeader())
                //{
                //    ;
                //}

                csvFileReader.Read();
                csvFileReader.ReadHeader();
                //var nameIndex = csvFileReader.Context.NamedIndexes;
                //var headers = csvFileReader.Context.HeaderRecord;
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

                

                //var distinctList = Filialen.Select(f => f).GroupBy(f => f.MyEqualityProperty).Select(grp => grp.First());

                //List<String> duplicates = Filialen.GroupBy(x => x)
                //             .Where(g => g.Count() > 1)
                //             .Select(g => g.Key)
                //             .ToList();



                //https://github.com/JoshClose/CsvHelper/issues/948
                //while (await csvFileReader.ReadAsync())
                //{
                //    i++;

                //}
                //var columns = csvFileReader.Context.ColumnCount;
            }
        }
    }
}
