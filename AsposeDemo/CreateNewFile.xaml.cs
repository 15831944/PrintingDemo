using AsposeDemo.implementations;
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
using static AsposeDemo.MainWindow;

namespace AsposeDemo
{
    /// <summary>
    /// Interaction logic for CreateNewFile.xaml
    /// </summary>
    public partial class CreateNewFile : Window
    {
        private IExcelProcessor _excelProcessor;
        private UsedLibrary library;
        private static string[] _FORMATS = new string[] { ".xls", ".xlsx" };

        public CreateNewFile(UsedLibrary library)
        {
            InitializeComponent();
            this.library = library;
            cmbPrinters.ItemsSource = getExcelInstance().GetAvailablePrinters();
            cmbFileType.ItemsSource = _FORMATS;
        }

        public IExcelProcessor getExcelInstance()
        {
            if(object.Equals(null, _excelProcessor))
            {
                if (this.library == UsedLibrary.Aspose)
                    _excelProcessor = new AsposeExcel();
                else if (this.library == UsedLibrary.Spire)
                    _excelProcessor = new SpireExcel();
                else if (this.library == UsedLibrary.GemboxSpreadSheet)
                    _excelProcessor = new GemboxSpreadSheetExcel();
            }

            return _excelProcessor;
        }

        private void cmbPrinters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            getExcelInstance().PrinterName = cmbPrinters.SelectedItem.ToString();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtFileName.Text.Trim()) || cmbFileType.SelectedIndex < 0 || cmbPrinters.SelectedIndex < 0)
            {
                MessageBox.Show(this, "You need to type a file name, select a type and select a printer in order to create a new file.", "Atention", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            getExcelInstance().FileName = string.Format(@"{0}\{1}{2}", System.IO.Directory.GetCurrentDirectory(), txtFileName.Text, cmbFileType.SelectedItem.ToString());
            getExcelInstance().CreateFile(this);
        }
    }
}
