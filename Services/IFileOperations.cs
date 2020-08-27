using DOCXAutomationWPF.Model;
using DOCXAutomationWPF.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Threading.Tasks;

namespace DOCXAutomationWPF.Services
{
    public interface IFileOperations
    {
        Task<HashSet<string>> ScanFile(string filename);
        //Task<ObservableCollection<ReplacementValues>> GetReplacementValuesAsync();
    }
}
