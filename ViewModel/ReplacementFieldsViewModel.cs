﻿using DOCXAutomationWPF.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Controls;
using Xceed.Words.NET;

namespace DOCXAutomationWPF.ViewModel
{
    public class ReplacementFieldsViewModel : ObservableCollection<ReplacementValues>
        //, INotifyPropertyChanged
    {
        public ReplacementFieldsViewModel() { }
        public ReplacementFieldsViewModel(HashSet<string> fields)
        {
            foreach (var field in fields)
            {
                //Add(new ReplacementValues(field));
            }
        }

        //private ObservableCollection<ReplacementValues> _replacementValues;
        //public ObservableCollection<ReplacementValues> ReplacementValues
        //{
        //    get => _replacementValues;
        //    set
        //    {
        //        _replacementValues = value;
        //        OnPropertyChanged();
        //    }
        //}

        //public event PropertyChangedEventHandler PropertyChanged;

        //protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        //{
        //    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        //}
    }
}
