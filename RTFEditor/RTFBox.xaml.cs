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

            RichTextControlWF.DisableBeep();
            CommandBindings.Add(new CommandBinding(SearchCommand, SearchCommandExecute));
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

        #region command
        public static RoutedUICommand SearchCommand = new RoutedUICommand("SearchCommand", "SearchCommand", typeof(RTFBox));
        private void SearchCommandExecute(object sender, ExecutedRoutedEventArgs e)
        {
            txtSearch.Focus();
        }
        #endregion

        #region ControlHandler

        //
        // Nach dem Laden des Control
        //
        private void RTFEditor_Loaded(object sender, RoutedEventArgs e)
        {
            // Schrifttypen- und -größen-Initialisierung
            Fonttype.SelectedValue = "ＭＳ ゴシック";
            Fontheight.SelectedValue = "12";

            DisableDefaultFont();

            // aktuelle Zeilen- und Spaltenpositionen angeben
            aktZeile = Zeilennummer();
            aktSpalte = Spaltennummer();           
        }


        #endregion ControlHandler

        #region private ToolBarHandler

        private void ToolStripButtonBold_Click(object sender, RoutedEventArgs e)
        {
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null) {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(font, font.Style ^ System.Drawing.FontStyle.Bold);
                RichTextControlWF.Focus();
            }
        }

        private void ToolStripButtonItalic_Click(object sender, RoutedEventArgs e)
        {
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null)
            {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(font, font.Style ^ System.Drawing.FontStyle.Italic);
                RichTextControlWF.Focus();
            }
        }

        private void ToolStripButtonUnderline_Click(object sender, RoutedEventArgs e)
        {
            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null)
            {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(font, font.Style ^ System.Drawing.FontStyle.Underline);
                RichTextControlWF.Focus();
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
                RichTextControlWF.Focus();
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
            RichTextControlWF.Focus();
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
            RichTextControlWF.Focus();
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

            if (fontHeight == null) {
                fontHeight = "12";
            }

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

        //
        // andere Font wurde vom Benutzer ausgewählt
        //
        private void Fonttype_DropDownClosed(object sender, EventArgs e)
        {            
            string fontName = (string)Fonttype.SelectedItem;

            if (fontName == null) { 
                fontName = "ＭＳ ゴシック";
            }

            System.Drawing.Font font = RichTextControlWF.SelectionFont;
            if (font != null)
            {
                RichTextControlWF.SelectionFont = new System.Drawing.Font(fontName, font.SizeInPoints, font.Style);
            }
            else {
                string fontHeight = (string)Fontheight.SelectedItem;
                if (fontHeight == null) fontHeight = "12";
                RichTextControlWF.SelectionFont = new System.Drawing.Font(fontName, float.Parse(fontHeight));
            }
            RichTextControlWF.Focus();
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
                RichTextControlWF.Focus();
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
                RichTextControlWF.Focus();
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
                RichTextControlWF.Focus();
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
            RichTextControlWF.Focus();

        }

        private void ToolStripButtonIndent_Click(object sender, RoutedEventArgs e)
        {
            int indent_size = 16;

            if (RichTextControlWF.SelectionFont != null)
            {
                indent_size = RichTextControlWF.SelectionFont.Height * 2;
            }

            RichTextControlWF.SelectionIndent += indent_size;
            RichTextControlWF.Focus();
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
            RichTextControlWF.Focus();

        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            mFindPos = 0;
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            ToolStripButtonSearch_Click(null, null);
        }

        int mFindPos = 0;
        private void ToolStripButtonSearch_Click(object sender, RoutedEventArgs e)
        {
            if (txtSearch.Text.Length == 0) return;

            //文字列を検索する
            mFindPos = RichTextControlWF.Find(txtSearch.Text, mFindPos, System.Windows.Forms.RichTextBoxFinds.MatchCase);

            //指定文字列が見つかったか？
            if (mFindPos > -1)
            {
                //見つかった位置から、文字数分を選択
                RichTextControlWF.Select(mFindPos, txtSearch.Text.Length);
                //フォーカスを当てる（フォーカスがはずれると選択が無効になってしまうため）
                RichTextControlWF.Focus();
                mFindPos += txtSearch.Text.Length;
            }
            else
            {
                mFindPos = 0;
            }
        }

        private System.Drawing.Image LoadImageWithSize(string sourceFile, int width, int height)
        {
            // サイズ変更する画像ファイルを開く
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(sourceFile))
            {
                // 変更倍率を取得する
                float scale = Math.Min((float)width / (float)image.Width, (float)height / (float)image.Height);

                // サイズ変更した画像を作成する
                System.Drawing.Bitmap bitmap = null;
                // 変更サイズを取得する
                int widthToScale = (int)(image.Width * scale);
                int heightToScale = (int)(image.Height * scale);

                bitmap = new System.Drawing.Bitmap(widthToScale, heightToScale);
                using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(bitmap)) { 

                    // 背景色を塗る
                    System.Drawing.SolidBrush solidBrush = new System.Drawing.SolidBrush(System.Drawing.Color.White);
                    graphics.FillRectangle(solidBrush, new System.Drawing.RectangleF(0, 0, width, height));

                    // サイズ変更した画像に、左上を起点に変更する画像を描画する
                    graphics.DrawImage(image, 0, 0, widthToScale, heightToScale);

                }

                return bitmap;
            }
        }

        private void ToolStripButtonInsertImage_Click(object sender, RoutedEventArgs e) {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = ""; // Default file name
            dlg.Filter = "All Image Files|*.bmp;*.ico;*.gif;*.jpeg;*.jpg;*.png;*.tif;*.tiff|" +
                             "Windows Bitmap(*.bmp)|*.bmp|" +
                             "Windows Icon(*.ico)|*.ico|" +
                             "Graphics Interchange Format (*.gif)|(*.gif)|" +
                             "JPEG File Interchange Format (*.jpg)|*.jpg;*.jpeg|" +
                             "Portable Network Graphics (*.png)|*.png|" +
                             "Tag Image File Format (*.tif)|*.tif;*.tiff";

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                try
                {
                    // If file is an icon
                    if (dlg.FileName.ToUpper().EndsWith(".ICO"))
                    {
                        // Create a new icon, get it's handle, and create a bitmap from
                        // its handle
                        RichTextControlWF.InsertImage(System.Drawing.Bitmap.FromHicon((new System.Drawing.Icon(dlg.FileName)).Handle));
                        //System.Windows.Forms.Clipboard.SetImage(System.Drawing.Bitmap.FromHicon((new System.Drawing.Icon(dlg.FileName)).Handle));
                        //RichTextControlWF.Paste();
                    }
                    else
                    {
                        // Create a bitmap from the filename
                        System.Drawing.Bitmap bmp = (System.Drawing.Bitmap)LoadImageWithSize(dlg.FileName, 640, 640);
                        if (bmp != null) {
                            RichTextControlWF.InsertImage(bmp);
                            //System.Windows.Forms.Clipboard.SetImage(bmp);
                            //RichTextControlWF.Paste();
                        }
                        bmp.Dispose();
                    }
                }
                catch (Exception _e)
                {
                    MessageBox.Show("Error : " + _e.GetType().Name + "\nDeatail : " + _e.Message + "\nStack : " + _e.StackTrace);

                }
            }
        }

        private void ToolStripButtonInsertTable_Click(object sender, RoutedEventArgs e) {
            DialogCreatTable dlg = new DialogCreatTable();
            dlg.Owner = Window.GetWindow(this);
            if (dlg.ShowDialog() == true) {
                RichTextControlWF.InsertTable(dlg.row, dlg.column, 100);
            }
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
                Fonttype.SelectedValue = "ＭＳ ゴシック";//System.Windows.Forms.Control.DefaultFont.Name;
                Fontheight.SelectedValue = "12";
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

        const uint EM_LINEINDEX = 0xBB;
        const uint EM_LINEFROMCHAR = 0xC9;
        const uint IMF_DUALFONT = 0x80;
        const uint WM_USER = 0x0400;
        const uint EM_SETLANGOPTIONS = WM_USER + (uint)120;
        const uint EM_GETLANGOPTIONS = WM_USER + (uint)121;

        [DllImport("User32.Dll")]
        private static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

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

        //デフォルトフォントを無効にする
        private void DisableDefaultFont() {
            uint lParam;
            lParam = (uint)SendMessage(RichTextControlWF.Handle, EM_GETLANGOPTIONS, 0, 0);
            lParam &= ~IMF_DUALFONT;
            SendMessage(RichTextControlWF.Handle, EM_SETLANGOPTIONS, 0, (int)lParam);
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
            System.Drawing.Font font = GetCurrentFont();

            RichTextControlWF.Rtf = rtf;
            if (rtf == "") {
                RichTextControlWF.Font = font;
                RichTextControlWF.SelectionFont = font;
            }
            RichTextControlWF.ZoomFactor = (float)SliderZoom.Value;

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
                    RichTextControlWF.ZoomFactor = (float)SliderZoom.Value;
                }
            }
        }

        //
        // ファイルダイアログを表示し、RTFファイルをロードする
        //
        public void SaveRtfFile(String name)
        {
            // Configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = name; // Default file name
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


        //現在のフォントを取得する
        private System.Drawing.Font GetCurrentFont()
        {
            string fontName = (string)Fonttype.SelectedItem;

            if (fontName == null) fontName = "ＭＳ ゴシック";
            string fontHeight = (string)Fontheight.SelectedItem;
            if (fontHeight == null) fontHeight = "12";
            return new System.Drawing.Font(fontName, float.Parse(fontHeight));
        }

        #endregion öffentliche Funktionen

        private void RichTextControlWF_LinkClicked(object sender, System.Windows.Forms.LinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(e.LinkText);
            }
            catch (Exception ex) {
                MessageBox.Show("エラー", ex.Message);
            }
        }

        private void RTFEditor_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key) {
            case Key.F3:
                ToolStripButtonSearch_Click(null, null);
                e.Handled = true;
                break;
            case Key.Escape:
                RichTextControlWF.Focus();
                e.Handled = true;
                break;
            }
        }
    }
}