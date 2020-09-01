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

        private async void ButtStartImport_Click(object sender, RoutedEventArgs e)
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

                //if (csvFileReader.ReadHeader())
                //{
                //    ;
                //}

                csvFileReader.Read();
                csvFileReader.ReadHeader();
                var nameIndex = csvFileReader.Context.NamedIndexes;
                var headers = csvFileReader.Context.HeaderRecord;
                var recList = csvFileReader.GetRecords<dynamic>().ToList();
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
