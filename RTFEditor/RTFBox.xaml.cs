using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Runtime.InteropServices;
using CustomWPFColorPicker;
using System.Diagnostics;

namespace RTFEditor
{
	/// <summary>
	/// Interaktionslogik für "RTFBox.xaml"
	/// </summary>
	public partial class RTFBox 
    {
        /// <summary>
        /// Konstruktor - initialisiert alle graphischen Komponenten
        /// </summary>
        public RTFBox()
        {
            this.InitializeComponent();

            //textColorPicker.ColorSelected += TextColorSelectedHandler;
            //backColorPicker.ColorSelected += BackColorSelectedHandler;
        }
       
        #region Variablen

        private bool dataChanged = false; // ungespeicherte Textänderungen     

        private string privateText = null; // Inhalt der RTFBox im txt-Format
        public string text
        {
            get
            {
                //TextRange range = new TextRange(RichTextControl.Document.ContentStart, RichTextControl.Document.ContentEnd);
                //return range.Text;
               return RichTextControlWF.Text;
            }
            set
            {
                privateText = value;
            }
        }

        private int privAktZeile = 1; 
        public int aktZeile
        {
            get { return privAktZeile; }
            set 
            { 
                privAktZeile = value;
                UpdateCurolLocation(privAktZeile, privAktSpalte);
            }
        }

        private int privAktSpalte = 1; 
        public int aktSpalte
        {
            get { return privAktSpalte; }
            set 
            { 
                privAktSpalte = value;
                UpdateCurolLocation(privAktZeile, privAktSpalte);
            }
        }
       
        #endregion Variablen     

        #region ControlHandler

        //
        // Nach dem Laden des Control
        //
        private void RTFEditor_Loaded(object sender, RoutedEventArgs e)
        {
            // Schrifttypen- und -größen-Initialisierung
            Fonttype.SelectedValue = System.Windows.Forms.Control.DefaultFont.Name;
            Fontheight.SelectedValue = "10";

            // aktuelle Zeilen- und Spaltenpositionen angeben
            aktZeile = Zeilennummer();
            aktSpalte = Spaltennummer();           

            RichTextControlWF.Focus();
        }


        #endregion ControlHandler

        #region private ToolBarHandler

        private void ToolStripButtonBold_Click(object sender, RoutedEventArgs e)
        {
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null) {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(font, font.Style ^ System.Drawing.FontStyle.Bold);
            }
        }

        private void ToolStripButtonItalic_Click(object sender, RoutedEventArgs e)
        {
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null)
            {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(font, font.Style ^ System.Drawing.FontStyle.Italic);
            }
        }

        private void ToolStripButtonUnderline_Click(object sender, RoutedEventArgs e)
        {
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null)
            {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(font, font.Style ^ System.Drawing.FontStyle.Underline);
            }
        }

        //
        // ToolStripButton Strikeout wurde gedrückt
        // (läßt sich nicht durch Command in XAML abarbeiten)
        //
        private void ToolStripButtonStrikeout_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null)
            {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(font, font.Style ^ System.Drawing.FontStyle.Strikeout);
            }
        }

        //
        // ToolStripButton Textcolor wurde gedrückt
        //
        private void TextColorSelectedHandler(object sender, RoutedEventArgs e)
        {
            //TextRange range = new TextRange(RichTextControl.Selection.Start, RichTextControl.Selection.End);
            //range.ApplyPropertyValue(FlowDocument.ForegroundProperty, new SolidColorBrush(textColorPicker.CurrentColor.Color));
            RichTextControlWF.SelectionColor = System.Drawing.Color.FromArgb(
                textColorPicker.CurrentColor.Color.A, 
                textColorPicker.CurrentColor.Color.R, 
                textColorPicker.CurrentColor.Color.G, 
                textColorPicker.CurrentColor.Color.B);
        }

        //
        // ToolStripButton Backgroundcolor wurde gedrückt
        //
        private void BackColorSelectedHandler(object sender, RoutedEventArgs e) {

            //TextRange range = new TextRange(RichTextControl.Selection.Start, RichTextControl.Selection.End);
            //range.ApplyPropertyValue(FlowDocument.BackgroundProperty, new SolidColorBrush(backColorPicker.CurrentColor.Color));
            RichTextControlWF.SelectionBackColor = System.Drawing.Color.FromArgb(
                backColorPicker.CurrentColor.Color.A,
                backColorPicker.CurrentColor.Color.R,
                backColorPicker.CurrentColor.Color.G,
                backColorPicker.CurrentColor.Color.B);
        }

        //
        // ToolStripButton Subscript wurde gedrückt
        // (läßt sich auch durch Command in XAML machen. 
        // Damit dann aber ein wirkliches Subscript angezeigt wird, muß die benutzte Font OpenType sein:
        // http://msdn.microsoft.com/en-us/library/ms745109.aspx#variants
        // Um auch für alle anderen Schriftarten Subscript zu realisieren, läßt sich stattdessen die Baseline Eigenschaft verändern)
        //

        //
        // Textgröße wurde vom Benutzer verändert
        //
        private void Fontheight_DropDownClosed(object sender, EventArgs e)
        {
            string fontHeight = (string)Fontheight.SelectedItem;

            if (fontHeight != null)
            {
                //RichTextControl.Selection.ApplyPropertyValue(System.Windows.Controls.RichTextBox.FontSizeProperty, fontHeight);

                System.Drawing.Font font = RichTextControlWF.SelectionFont;
                if (font != null)
                {
                    RichTextControlWF.SelectionFont = new System.Drawing.Font(font.FontFamily, float.Parse(fontHeight), font.Style);
                }
                else {
                    RichTextControlWF.SelectionFont = new System.Drawing.Font(System.Windows.Forms.Control.DefaultFont.FontFamily, float.Parse(fontHeight));
                }
                RichTextControlWF.Focus();
            }
        }

        //
        // andere Font wurde vom Benutzer ausgewählt
        //
        private void Fonttype_DropDownClosed(object sender, EventArgs e)
        {            
            string fontName = (string)Fonttype.SelectedItem;

            if (fontName != null)
            {
                System.Drawing.Font font = RichTextControlWF.SelectionFont;
                if (font != null)
                {
                    RichTextControlWF.SelectionFont = new System.Drawing.Font(fontName, font.SizeInPoints, font.Style);
                }
                else {
                    string fontHeight = (string)Fontheight.SelectedItem;
                    if (fontHeight == null) fontHeight = "10";
                    RichTextControlWF.SelectionFont = new System.Drawing.Font(fontName, float.Parse(fontHeight));
                }
                RichTextControlWF.Focus();
            }
        }

        //
        // Alignmentstatus anpassen
        //
        private void ToolStripButtonAlignLeft_Click(object sender, RoutedEventArgs e)
        {
            if (ToolStripButtonAlignLeft.IsChecked == true)
            {
                ToolStripButtonAlignCenter.IsChecked = false;
                ToolStripButtonAlignRight.IsChecked = false;

                RichTextControlWF.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Left;
            }
        }

        //
        // Alignmentstatus anpassen
        //
        private void ToolStripButtonAlignCenter_Click(object sender, RoutedEventArgs e)
        {
            if (ToolStripButtonAlignCenter.IsChecked == true)
            {
                ToolStripButtonAlignLeft.IsChecked = false;
                ToolStripButtonAlignRight.IsChecked = false;

                RichTextControlWF.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Center;
            }

        }

        //
        // Alignmentstatus anpassen
        //
        private void ToolStripButtonAlignRight_Click(object sender, RoutedEventArgs e)
        {
            if (ToolStripButtonAlignRight.IsChecked == true)
            {
                ToolStripButtonAlignCenter.IsChecked = false;
                ToolStripButtonAlignLeft.IsChecked = false;

                RichTextControlWF.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Right;
            }
        }

        private void SliderZoom_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (RichTextControlWF == null) return;
            RichTextControlWF.ZoomFactor = (float)SliderZoom.Value;
        }

        private void ToolStripButtonBulletList_Click(object sender, RoutedEventArgs e)
        {
            RichTextControlWF.SelectionBullet = !RichTextControlWF.SelectionBullet;

        }

        private void ToolStripButtonIndent_Click(object sender, RoutedEventArgs e)
        {
            int indent_size = 16;

            if (RichTextControlWF.SelectionFont != null)
            {
                indent_size = RichTextControlWF.SelectionFont.Height * 2;
            }

            RichTextControlWF.SelectionIndent += indent_size;
        }

        private void ToolStripButtonIndentDelete_Click(object sender, RoutedEventArgs e)
        {
            int indent_size = 16;

            if (RichTextControlWF.SelectionFont != null)
            {
                indent_size = RichTextControlWF.SelectionFont.Height * 2;
            }

            RichTextControlWF.SelectionIndent -= indent_size;
            if (RichTextControlWF.SelectionIndent < 0) RichTextControlWF.SelectionIndent = 0;

        }
        #endregion private ToolBarHandler

        #region private RichTextBoxHandler

        private void RichTextControlWF_SelectionChanged(object sender, EventArgs e)
        {
            // markierten Text holen
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            
            
            ToolStripButtonBold.IsChecked = (font != null && font.Bold);
            ToolStripButtonItalic.IsChecked = (font != null && font.Italic);
            ToolStripButtonUnderline.IsChecked = (font != null && font.Underline);
            ToolStripButtonStrikeout.IsChecked = (font != null && font.Strikeout);

            ToolStripButtonAlignLeft.IsChecked = RichTextControlWF.SelectionAlignment == System.Windows.Forms.HorizontalAlignment.Left;
            ToolStripButtonAlignCenter.IsChecked = RichTextControlWF.SelectionAlignment == System.Windows.Forms.HorizontalAlignment.Center;
            ToolStripButtonAlignRight.IsChecked = RichTextControlWF.SelectionAlignment == System.Windows.Forms.HorizontalAlignment.Right;

            ToolStripButtonBulletList.IsChecked = RichTextControlWF.SelectionBullet;

            // Schrifttypen- und -größen-Initialisierung
            if (RichTextControlWF.SelectionFont != null) {
                if(RichTextControlWF.SelectionFont.Name != null)
                {
                    Fonttype.SelectedValue = RichTextControlWF.SelectionFont.Name;

                }
                Fontheight.SelectedValue = RichTextControlWF.SelectionFont.SizeInPoints.ToString();
            }
            else
            {
                Fonttype.SelectedValue = System.Windows.Forms.Control.DefaultFont.Name;
                Fontheight.SelectedValue = "10";
            }

            // Ausgabe der Zeilennummer
            aktZeile = Zeilennummer();

            // Ausgabe der Spaltennummer
            aktSpalte = Spaltennummer();
        }

        //
        // wurden Textänderungen gemacht
        //
        private void RichTextControlWF_TextChanged(object sender, EventArgs e)
        {
            dataChanged = true;
        }

        //
        // Tastendruck erzeugt ein neues Zeichen in der gewählten Font
        //
        private void RichTextControlWF_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            dataChanged = true;
        }

        //
        // Tastenkombinationen auswerten
        //
        private void RichTextControlWF_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            // Ctrl + B
            if ((Keyboard.Modifiers == ModifierKeys.Control) && (e.KeyCode == System.Windows.Forms.Keys.B))
            {
                ToolStripButtonBold.IsChecked = !ToolStripButtonBold.IsChecked;
            }

            // Ctrl + I
            if ((Keyboard.Modifiers == ModifierKeys.Control) && (e.KeyCode == System.Windows.Forms.Keys.I))
            {
                ToolStripButtonItalic.IsChecked = !ToolStripButtonItalic.IsChecked;
            }

            // Ctrl + U
            if ((Keyboard.Modifiers == ModifierKeys.Control) && (e.KeyCode == System.Windows.Forms.Keys.U))
            {
                ToolStripButtonUnderline.IsChecked = !ToolStripButtonUnderline.IsChecked;
            }
        }

        #endregion private RichTextBoxHandler

        #region private Funktionen

        const int EM_LINEINDEX = 0xBB;
        const int EM_LINEFROMCHAR = 0xC9;

        [DllImport("User32.Dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        //
        // Gibt die Zeilennummer der aktuellen Cursorposition zurück
        //
        private int Zeilennummer()
        {
            return SendMessage(RichTextControlWF.Handle, EM_LINEFROMCHAR, -1, 0) + 1;
        }

        //
        // Gibt die Spaltennummer der aktuellen Cursorposition zurück
        private int Spaltennummer()
        {
            int lineIndex = SendMessage(RichTextControlWF.Handle, EM_LINEINDEX, -1, 0);
            return RichTextControlWF.SelectionStart - lineIndex + 1;
        }

        //ステータスバーのカーソル位置を更新する
        private void UpdateCurolLocation(int row, int columun) {
            LabelZeileSpalte.Text = String.Format("行:{0} 列:{1}", row,  columun);
        }

        #endregion private Funktionen

        # region öffentliche Funktionen

        //
        // Alle Daten löschen
        //
        public void Clear()
        {            
            dataChanged = false;            
            RichTextControlWF.Clear();            
        }

        //
        // Inhalt der RichTextBox als RTF setzen
        //
        public void SetRTF(string rtf)
        {
            RichTextControlWF.Rtf = rtf;
        }

        //
        // RTF Inhalt der RichTextBox als RTF-String zurückgeben
        //
        public string GetRTF()
        {
            return RichTextControlWF.Rtf;
        }

        //データの変更の有無を取得する
        public bool IsDataChenged() {
            return dataChanged;
        }

        //
        // ファイルダイアログを表示し、RTFファイルをロードする
        //
        public void LoadRtfFile()
        {
            SetRTF("");


            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = ""; // Default file name
            dlg.DefaultExt = ".rtf"; // Default file extension
            dlg.Filter = "RichText documents (.rtf)|*.rtf|Text documents (.txt)|*.txt"; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                string path = dlg.FileName;

                if (File.Exists(path))
                {
                    RichTextControlWF.LoadFile(path);
                }
            }
        }

        //
        // ファイルダイアログを表示し、RTFファイルをロードする
        //
        public void SaveRtfFile()
        {
            // Configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = ""; // Default file name
            dlg.DefaultExt = ".rtf"; // Default file extension
            dlg.Filter = "RichText documents (.rtf)|*.rtf"; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                string path = System.IO.Path.ChangeExtension(dlg.FileName, "rtf");

                // Open document
                RichTextControlWF.SaveFile(path);
            }
        }

        //
        // 現在のドキュメントを印刷する
        //
        public void PrintDocument()
        {

            System.Windows.Forms.PrintDialog dlg = new System.Windows.Forms.PrintDialog();
            dlg.Document = new System.Drawing.Printing.PrintDocument();

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                RichTextControlWF.Print(dlg.Document);
            }
        }

        #endregion öffentliche Funktionen

    }
}