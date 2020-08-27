using DOCXAutomationWPF.Model;
using DOCXAutomationWPF.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Xceed.Words.NET;

namespace DOCXAutomationWPF.Services
{
    public class FileOperations : INotifyPropertyChanged
    {
        public HashSet<string> replacementTags;
        private string _filename;
        public string Filename
        {
            get => _filename;
            set { _filename = value; NotifyPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        //public Task<ObservableCollection<ReplacementValues>> GetReplacementValuesAsync()
        //{
        //    return Task.FromResult();
        //}

        public Task<HashSet<string>> ScanFile(string filename)
        {
            replacementTags = new HashSet<string>();

            using (var doc = DocX.Load(filename))
            {
                var matches = new List<string>();
                if ((matches = doc.FindUniqueByPattern(@"<[\w \=]{4,}>", RegexOptions.IgnoreCase)).Count > 0)
                {
                    foreach (var match in matches)
                        replacementTags.Add(match);
                }
            }

            return Task.FromResult(replacementTags);
        }
    }
}
