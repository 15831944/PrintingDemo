using AsposeDemo.implementations;
using System;
using System.Windows;
using System.Windows.Controls;
using static AsposeDemo.MainWindow;

namespace AsposeDemo
{
    /// <summary>
    /// Interaction logic for PrintExistingFile.xaml
    /// </summary>
    public partial class PrintExistingFile : Window
    {
        private IExcelProcessor _excelProcessor;
        private UsedLibrary library;

        public PrintExistingFile(UsedLibrary library)
        {
            InitializeComponent();
            this.library = library;
            if(library == UsedLibrary.Spire)
            {
                lblPrinter.Visibility = Visibility.Hidden;
                cmbPrinters.Visibility = Visibility.Hidden;
            }
            cmbPrinters.ItemsSource = getExcelInstance().GetAvailablePrinters();
        }

        public IExcelProcessor getExcelInstance()
        {
            if (object.Equals(null, _excelProcessor))
            {
                if (library == UsedLibrary.Aspose)
                    this._excelProcessor = new AsposeExcel();
                else if (library == UsedLibrary.Spire)
                    this._excelProcessor = new SpireExcel();
                else if (library == UsedLibrary.GemboxSpreadSheet)
                    this._excelProcessor = new GemboxSpreadSheetExcel();

            }

            return _excelProcessor;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(getExcelInstance().FileName))
            {
                MessageBox.Show(this, "You need to choose a file in order to print it.", "Atention", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            getExcelInstance().PrintFile(this, string.Empty);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls";

            Nullable<bool> result = dlg.ShowDialog();

            if (result.HasValue && result.Value)            
                lblBrowsedFile.Content = getExcelInstance().FileName = dlg.FileName;            
            else
                lblBrowsedFile.Content = getExcelInstance().FileName = string.Empty;

            getExcelInstance().PrinterName = string.Empty;
            cmbPrinters.SelectedIndex = -1;
        }

        private void cmbPrinters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cmbPrinters.SelectedIndex >= 0)
                getExcelInstance().PrinterName = cmbPrinters.SelectedItem.ToString();
        }
    }
}
