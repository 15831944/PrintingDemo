using System;
using System.Collections.Generic;
using System.Data;
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
    /// Interaction logic for TimesOfProcessing.xaml
    /// </summary>
    public partial class TimesOfProcessing : Window
    {
        public TimesOfProcessing(Dictionary<string, List<string>> values)
        {
            InitializeComponent();

            dtgTimes.ItemsSource = ProcessingData(values).AsDataView();
        }
        private DataTable ProcessingData(Dictionary<string, List<string>> values)
        {
            DataTable _dt = new DataTable();

            _dt.Columns.Add("Format", Type.GetType("System.String"));
            _dt.Columns.Add("TimeElapsed", Type.GetType("System.String"));

            foreach (var key in values.Keys)
            {
                List<string> listOfValues = null;
                values.TryGetValue(key, out listOfValues);

                foreach(var value in listOfValues)
                    _dt.Rows.Add(new Object[] { key, value });
            }

            return _dt;
        }
    }
}
