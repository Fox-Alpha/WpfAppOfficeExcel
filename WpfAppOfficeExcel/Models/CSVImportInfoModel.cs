using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfAppOfficeExcel.Models
{
    public class CSVImportInfoModel : INotifyPropertyChanged
    {
        private int _anzahlFilialen;

        public int AnzahlFiliale
        {
            get { return _anzahlFilialen; }
            set { _anzahlFilialen = value; OnPropertyRaised("AnzahlFilialen"); }
        }

        private string _importFileName;

        public string ImportFileName
        {
            get { return _importFileName; }
            set { _importFileName = value; OnPropertyRaised("ImportFileName"); }
        }

        private string _ExportFileName;

        public string ExportFileName
        {
            get { return _ExportFileName; }
            set { _ExportFileName = value; OnPropertyRaised("ExportFileName"); }
        }

        public string ImportFilename { get; }

        public CSVImportInfoModel(string importFilename, string exportFileName)
        {
            ImportFileName = importFilename;
            ExportFileName = exportFileName;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyRaised(string propName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }
    }
}
