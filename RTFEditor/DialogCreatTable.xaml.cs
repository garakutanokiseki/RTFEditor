using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace RTFEditor
{
    /// <summary>
    /// DialogCreatTable.xaml の相互作用ロジック
    /// </summary>
    public partial class DialogCreatTable : Window
    {
        public int row = 2;
        public int column = 2;

        public DialogCreatTable()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;

            try
            {
                column = int.Parse(textColumn.Text);
                row = int.Parse(textRow.Text);
            }
            catch (Exception ex) {
                row = 2;
                column = 2;
            }

            Close();

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();

        }
    }
}
