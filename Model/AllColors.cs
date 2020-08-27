using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;

namespace DOCXAutomationWPF.Model
{
    public class AllColors : List<Color>
    {
        public AllColors()
        {
            //foreach (var colorValue in Enum.GetValues(typeof(KnownColor)))
            //{
            //    Add(Color.FromKnownColor((KnownColor)colorValue));
            //}

            foreach (PropertyInfo property in typeof(Color).GetProperties())
                if (property.PropertyType == typeof(Color))
                    Add((Color)property.GetValue(null));            
        }
    }
}
