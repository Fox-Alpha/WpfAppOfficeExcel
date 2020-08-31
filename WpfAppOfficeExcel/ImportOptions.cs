﻿using System;
using System.ComponentModel;

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


        private enumImportOptions activeImportOptions = enumImportOptions.None;
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
                    if ((activeImportOptions & value) == value)
                    {
                        activeImportOptions &= ~value;
                    }
                    else
                        activeImportOptions |= value;

                    OnPropertyRaised("strImpOpt");
                }
            }
        }

        public string strImpOpt
        {
            get
            {
                return ActiveImportOptions.ToString();
            }
        }

        public ImportOptions()
        {
            //ActiveImportOptions = enumImportOptions.None;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        //public void NotifyPropertyChanged(string propName)
        
        public void OnPropertyRaised(string propName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }
    }
}