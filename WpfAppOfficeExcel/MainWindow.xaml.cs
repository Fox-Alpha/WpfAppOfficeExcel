using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
            }
        }
    }
}
