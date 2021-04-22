using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Documents;

namespace WpfAppOfficeExcel.Importer
{
        [Flags]
        public enum enumImportOptions
        {
            None = 0,
            WarenEingang = 1,
            WarenAusgang = 2,
            ProduktVerlauf = 4,
            ProduktRetoure = 8,
            RabattKunde = 16,
            WarenbewegungPositiv = 32,
            WarenbewegungNegativ = 64,
            UmlagerungEingang = 128,
            UmlagerungAusgang = 256,
            Inventur = 512
        }

    public class ImportOptions : INotifyPropertyChanged
    {
        public readonly Dictionary<string, string> dictImportOptions = new Dictionary<string, string>()
        {
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.None), ""},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.WarenEingang), "WE"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.WarenAusgang), "WA"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.ProduktVerlauf), "PV"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.ProduktRetoure), "PR"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.RabattKunde), "RN"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.WarenbewegungPositiv), "WP"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.WarenbewegungNegativ), "WN"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.UmlagerungEingang), "UE"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.UmlagerungAusgang), "UA"},
            {Enum.GetName(typeof(enumImportOptions), enumImportOptions.Inventur), "MI"},     // ???
        };

        private bool expKmpgColumns;

        public bool ExpKmpgColumns
        {
            get { return expKmpgColumns; }
            set
            {
                expKmpgColumns = value;
                OnPropertyRaised("ExpKmpgColumns");
            }
        }

        private bool oneSheetOnly = false;
        public bool OneSheetOnly { get => oneSheetOnly; set { oneSheetOnly = value; OnPropertyRaised("OneSheetOnly"); } }

        private enumImportOptions activeImportOptions = enumImportOptions.None;
        private bool oneSheetOnly1 = false;

        public enumImportOptions ActiveImportOptions
        {
            get { return activeImportOptions; }
            set
            {
                if (value == enumImportOptions.None)
                {
                    activeImportOptions = value;
                }
                else
                {
                    activeImportOptions ^= enumImportOptions.None;

                    if ((activeImportOptions & value) == value)
                    {
                        activeImportOptions &= ~value;
                    }
                    else
                        activeImportOptions |= value;

                    OnPropertyRaised("strImpOpt");
                }
                OnPropertyRaised("ActiveImportOptions");
            }
        }

        public string strImpOpt
        {
            get
            {
                return ActiveImportOptions.ToString();
            }
        }

        public List<string> GetImportOptionsAsList()
        {
            List<string> listImpOpt = new List<string>();

            var vals = Enum.GetValues(typeof(enumImportOptions));

            foreach (enumImportOptions val in vals)
            {
                if ((activeImportOptions & val) == val)
                {
                    if (dictImportOptions.TryGetValue(Enum.GetName(typeof(enumImportOptions), val), out string shortVal))
                    {
                        listImpOpt.Add(shortVal);
                    }
                }
            }

            return listImpOpt;
        }

        public ImportOptions()
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyRaised(string propName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }
    }
}
