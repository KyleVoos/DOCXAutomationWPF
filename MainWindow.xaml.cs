using DOCXAutomationWPF.Model;
using DOCXAutomationWPF.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace DOCXAutomationWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public List<Color> allColors = new AllColors().OrderBy(x => x.ToString()).ToList();
        private List<string> _fonts = System.Drawing.FontFamily.Families.Select(f => f.Name).ToList();
        public List<string> Fonts { get => _fonts; set => _fonts = value; }
        private FileOperations _fileOperations;
        private string _filename;
        public string Filename 
        {
            get => _filename;
            set => _filename = value;
        }
        public MainWindow()
        {
            InitializeComponent();
            _fileOperations = new FileOperations();
            _replacementValues = new ObservableCollection<ReplacementValues>();
            testBinding.ItemsSource = _replacementValues;
            Resources.Add("colors", allColors);
            Resources.Add("fontsList", Fonts);
        }

        private ObservableCollection<ReplacementValues> _replacementValues;
        public ObservableCollection<ReplacementValues> Data
        {
            get => _replacementValues;
            set
            {
                _replacementValues = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void openFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Docx files (*.docx)|*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _filename = openFileDialog.FileName;
                filenameLabel.Content = _filename;
            }
        }

        private void scanFIleBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_filename))
            {
                var vals = _fileOperations.ScanFile(_filename);
                AddValues(vals.Result);
            }
        }

        private void AddValues(HashSet<string> results)
        {
            foreach (var res in results)
                _replacementValues.Add(new ReplacementValues(res));
        }
    }
}
