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

namespace AsposeDemo
{
    /// <summary>
    /// Interaction logic for AsposeReport.xaml
    /// </summary>
    public partial class AsposeReport : Window
    {
        private IExcelProcessor _excelProcessor;
        private Dictionary<string, string> templates = new Dictionary<string, string>(3);

        public AsposeReport()
        {
            InitializeComponent();
            templates.Add("XLS", "AneSh1.xls");
            templates.Add("Excel SpreadSheet", "AneSh1.xlsx");
            templates.Add("Strict Open Xml", "AneSh1_soxs.xlsx");

            cmbPrinters.ItemsSource = getExcelInstance().GetAvailablePrinters();
            cmbTemplates.ItemsSource = templates.Keys;
        }

        public IExcelProcessor getExcelInstance()
        {
            if (object.Equals(null, _excelProcessor))            
                _excelProcessor = new AsposeExcel();            

            return _excelProcessor;
        }

        private void btnReport_Click(object sender, RoutedEventArgs e)
        {
            if(cmbTemplates.SelectedIndex < 0)
                MessageBox.Show(this, "You need to select a template before get the report.", "Atention", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            else 
                getExcelInstance().BuildReport(this);
        }

        private void cmbTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            getExcelInstance().FileName = templates[cmbTemplates.SelectedItem.ToString()];
        }

        private void cmbPrinters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            getExcelInstance().PrinterName = cmbPrinters.SelectedItem.ToString();
        }
    }
}
