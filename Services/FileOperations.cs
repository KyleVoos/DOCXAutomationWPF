using DOCXAutomationWPF.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace DOCXAutomationWPF.Services
{
    public class FileOperations : INotifyPropertyChanged
    {
        public HashSet<string> replacementTags;
        private string _filename;
        private MemoryStream memoryStream;
        public string Filename
        {
            get => _filename;
            set { _filename = value; NotifyPropertyChanged(); }
        }
        private DocX doc;
        private bool _isModified;
        public bool IsModified
        {
            get => _isModified;
            set { _isModified = value; NotifyPropertyChanged(); }
        }

        public FileOperations()
        {
            _isModified = false;
        }

        public FileOperations(string filename)
        {
            _isModified = false;
            _filename = filename;
        }        

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Scanning the .docx file selected by the user to find all the
        /// replacement tags and add each one to a HashSet to avoid duplicates.
        /// The values in the HashSet are then used to create an
        /// ObservableCollection<ReplacementValues>
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public Task<HashSet<string>> ScanFile(string filename)
        {
            _filename = filename;
            replacementTags = new HashSet<string>();

            using (doc = DocX.Load(filename))
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

        // Would be nice to implement a check on empty values before the user submits and give them a warning if any are found.
        /// <summary>
        /// After the user has entered in all the replacement values for the tags
        /// the document is opened again, each of the entries in the collection is 
        /// iterated through and the rags edited to the new value and formatting options.
        /// The edited file is saved to a MemoryStream because the file is not saved automatically.
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        public Task EditFile(ObservableCollection<ReplacementValues> values) 
        {
            try
            {
                using (doc = DocX.Load(Filename))
                {
                    foreach (var tag in values)
                    {
                        var bold = (tag.TextStyle == "Bold") ? true : false;
                        var italic = (tag.TextStyle == "Italic") ? true : false;
                        var underline = (tag.TextStyle == "Underline") ? true : false;
                        tag.ReplacementValue = (tag.ReplacementValue == null) ? "" : tag.ReplacementValue;
                        doc.ReplaceText(tag.ReplacementField, tag.ReplacementValue, false, RegexOptions.IgnoreCase, new Formatting()
                        {
                            FontColor = tag.TextColor,
                            FontFamily = new Font(tag.FontFamily),
                            Bold = bold,
                            Italic = italic,
                            Size = tag.FontSize
                        });
                    }

                    memoryStream = new MemoryStream();
                    doc.SaveAs(memoryStream);
                }

                _isModified = true;
            }
            catch (Exception ex)
            {
                return Task.FromException(ex);
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Saving the edited file. The file is opened from the memory stream where the file was previously
        /// saved after being edited.
        /// </summary>
        /// <returns></returns>
        public Task Save()
        {
            try
            {
                if (_isModified)
                {
                    using (doc = DocX.Load(memoryStream))
                        doc.SaveAs(_filename);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Task.FromException(ex);
            }

            return Task.CompletedTask;
        }

        public Task SaveAs(string saveFilePath)
        {
            try
            {
                if (_isModified)
                {
                    using (doc = DocX.Load(memoryStream))
                        doc.SaveAs(saveFilePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Task.FromException(ex);
            }

            return Task.CompletedTask;
        }
    }
}
