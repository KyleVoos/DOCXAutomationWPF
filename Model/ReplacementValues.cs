using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text;

namespace DOCXAutomationWPF.Model
{
    public class ReplacementValues : INotifyPropertyChanged
    {
        private string _replacementField;
        public string ReplacementField 
        {
            get => _replacementField;
            set
            {
                _replacementField = value;
                OnPropertyChanged();
            }
        }
        private string _replacementValue;
        public string ReplacementValue 
        {
            get => _replacementValue;
            set
            {
                _replacementValue = value;
                //OnPropertyChanged();
            }
        }
        private string _textStyle;
        public string TextStyle
        {
            get => _textStyle;
            set
            {
                _textStyle = value;
                OnPropertyChanged();
            }
        }
        private Color _textColor;
        public Color TextColor
        {
            get => _textColor;
            set
            {
                _textColor = value;
                OnPropertyChanged();
            }
        }
        private string _fontFamily;
        public string FontFamily
        {
            get => _fontFamily;
            set
            {
                _fontFamily = value;
                OnPropertyChanged();
            }
        }
        private int _fontSize = 11;
        public int FontSize
        {
            get => _fontSize;
            set
            {
                _fontSize = value;
                OnPropertyChanged();
            }
        }

        public ReplacementValues(string field)
        {
            _replacementField = field;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            //if (propertyName == "FontFamily")
            //{
            //    this.FontFamily = 
            //}
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
