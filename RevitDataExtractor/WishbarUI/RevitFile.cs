﻿using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RevitDataExtractor
{
    class RevitFile : INotifyPropertyChanged
    {
        private static Regex EndYear = new Regex(@"_20\d{2}$");
        private static Regex StartYear = new Regex(@"^20\d{2}_");
        public RevitFile()
        {

        }

        public RevitFile(string file)
        {
            Filename = Path.GetFileNameWithoutExtension(file);
            Extension = Path.GetExtension(file);
            Directory = Path.GetDirectoryName(file);
        }
        public string Filename { get; set; }
        private string _newFilename;
        public string NewFilename
        {
            set
            {
                _newFilename = value;
                OnPropertyChanged("NewFilename");
            }
            get
            {
                return _newFilename;
            }
        }
        public string Extension { get; set; }
        public string Directory { get; set; }
        public string Version { get; set; }
        public string AdditionalInfo { get; set; }
        public string FullPath
        {
            get
            {
                return Path.Combine(Directory, Filename + Extension);
            }
        }
        public string NewFullPath
        {
            get
            {
                return Path.Combine(Directory, NewFilename + Extension);
            }
        }

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion
    }
}
