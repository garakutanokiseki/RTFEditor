using System;
using System.Collections.ObjectModel;
using System.Windows.Media;


namespace RTFEditor
{
    /// <summary>
    /// Erzeugt eine String Liste aller auf dem System installierten Schriftarten
    /// </summary>
    class FontList : ObservableCollection<string> 
    { 
        public FontList() 
        {
            foreach (FontFamily f in Fonts.SystemFontFamilies)
            {
                System.Drawing.Font font = new System.Drawing.Font(f.ToString(), 8);
                this.Add(font.Name);                
            }  
        }   
    }
}
