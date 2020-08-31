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

namespace WpfAppOfficeExcel
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ImportOptions Import;

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
                gbSelectOptionForImport.IsEnabled = true;
                tbFilePathInfo.Text = openFileDialog.FileName;
            }
        }
    }

    public class ImportOptions : INotifyPropertyChanged
    {
        [Flags]
        public enum enumImportOptions
        {
            None,
            WarenEingang,
            WarenAusgang,
            ProduktVerlauf,
            ProduktRetoure,
            WarenbewegungPositiv,
            WarenbewegungNegativ,
            UmlagerungEingang,
            UmlagerungAusgang,
            Inventur
        }

        private enumImportOptions activeImportOptions;

        public enumImportOptions ActiveImportOptions
        {
            get { return activeImportOptions; }
            set 
            {
                if (activeImportOptions != enumImportOptions.None)
                {
                    return;
                    //activeImportOptions = value;
                }
                else
                {
                    activeImportOptions |= value;
                }                
            }
        }

        public ImportOptions()
        {
            ActiveImportOptions = enumImportOptions.None;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged(string propName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }
    }
}
