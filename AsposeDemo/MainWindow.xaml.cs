using System.Windows;

namespace AsposeDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();           
        }

        private void btnPrintFile_Click(object sender, RoutedEventArgs e)
        {
            if (isSelectedAnyLibrary())
            {
                PrintExistingFile pefWin = new PrintExistingFile(GetUsedLibray());
                pefWin.Owner = this;
                pefWin.Show();
            }
            else
                MessageBox.Show(this, "Please, select a library option before.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void btnCreateFile_Click(object sender, RoutedEventArgs e)
        {
            if (isSelectedAnyLibrary())
            {
                CreateNewFile cnfWin = new CreateNewFile(GetUsedLibray());
                cnfWin.Owner = this;
                cnfWin.Show();
            }
            else
                MessageBox.Show(this, "Please, select a library option before.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void btnGetReport_Click(object sender, RoutedEventArgs e)
        {
            if (isSelectedAnyLibrary())
            { 
                BuildReport cnfWin = new BuildReport(GetUsedLibray());
                cnfWin.Owner = this;
                cnfWin.Show();
            }
            else
                MessageBox.Show(this, "Please, select a library option before.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private UsedLibrary GetUsedLibray()
        {
            if (rbnAspose.IsChecked.Value)
                return UsedLibrary.Aspose;
            else if (rbnSpire.IsChecked.Value)
                return UsedLibrary.Spire;
            else if (rbnGembox.IsChecked.Value)
                return UsedLibrary.GemboxSpreadSheet;
            else
                return UsedLibrary.Other;
        }

        private bool isSelectedAnyLibrary()
        {
            return rbnAspose.IsChecked.Value || rbnSpire.IsChecked.Value || rbnGembox.IsChecked.Value;
        }

        public enum UsedLibrary
        {
            Aspose,
            Spire,
            GemboxSpreadSheet,
            Other
        }
    }
}
