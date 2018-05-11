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
    /// Interaction logic for BuildReport.xaml
    /// </summary>
    public partial class BuildReport : Window
    {
        private IExcelProcessor _excelProcessor;
        private UsedLibrary library;

        public BuildReport(UsedLibrary library)
        {
            InitializeComponent();
            this.library = library;
            cmbPrinters.ItemsSource = getExcelInstance().GetAvailablePrinters();
        }

        public IExcelProcessor getExcelInstance()
        {
            if (object.Equals(null, _excelProcessor))
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

        private void btnPrintReport_Click(object sender, RoutedEventArgs e)
        {
            if (cmbPrinters.SelectedIndex < 0)
            {
                MessageBox.Show(this, "You need to select a printer in order to print the report.", "Atention", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            
            getExcelInstance().BuildReport(this);
        }
    }
}
