using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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
        private Bitmap _excel { get; set; }

        public ExcelViewer(ImageSource excel)
        {
            InitializeComponent();
            this.excel = excel;
        }

        public ExcelViewer(Bitmap _excel)
        {
            InitializeComponent();
            this._excel = _excel;
        }

        private void Image_Loaded(object sender, RoutedEventArgs e)
        {                       
            var image = sender as System.Windows.Controls.Image;
            image.Source = !object.Equals(null, excel) ? excel : BitmapToImageSource(_excel);            
        }

        BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }
    }
}
