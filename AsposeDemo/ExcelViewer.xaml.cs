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
    /// Interaction logic for ExcelViewer.xaml
    /// </summary>
    public partial class ExcelViewer : Window
    {
        private ImageSource excel { get; set; }
        public ExcelViewer(ImageSource excel)
        {
            InitializeComponent();
            this.excel = excel;
        }

        private void Image_Loaded(object sender, RoutedEventArgs e)
        {
            // ... Get Image reference from sender.
            var image = sender as Image;
            // ... Assign Source.
            image.Source = excel;
        }
    }
}
