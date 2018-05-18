using Spire.Xls;
using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;

namespace AsposeDemo.implementations
{
    public class SpireExcel : IExcelProcessor
    {
        public string FileName { get; set; }
        public string FileFormatExtension { get; set; }
        public string PrinterName { get; set; }

        public void BuildReport(Window owner)
        {
            if (string.IsNullOrEmpty(PrinterName) && string.IsNullOrWhiteSpace(PrinterName))
            {
                MessageBox.Show(owner, "Please select your printer name.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                string dataDir = System.IO.Directory.GetCurrentDirectory();
                string fullFileName = dataDir + @"\templates\";

                Workbook workbook = new Workbook();
                workbook.LoadFromFile(fullFileName + "Template.xlsx");

                //Edit Text
                Worksheet sheet = workbook.Worksheets[0];
                DataTable data = GetData();
                ExcelFont fontBold = workbook.CreateFont();

                sheet.PageSetup.Orientation = PageOrientationType.Landscape;
                fontBold.IsBold = true;

                var headers = data.AsEnumerable()
                            .Select(s => s.Field<string>("Name"))
                            .ToArray();

                var cellsData = data.AsEnumerable()
                            .Select(s => s.Field<int>("Type"))
                            .ToArray();

                int columnsRange = 3 + headers.Length;

                sheet.InsertArray(headers, 16, 3, false);
                sheet.InsertArray(cellsData, 17, 3, false);

                CellRange range = sheet.Range[16, 3, 16, columnsRange];

                range.Style.Rotation = 60;
                range.Style.Font.IsBold = true;
                range.AutoFitRows();
                range.ColumnWidth = 4;

                sheet.Range[16, 3, 17, columnsRange].BorderInside(LineStyleType.Thin, Color.Black);
                sheet.Range[16, 3, 17, columnsRange].BorderAround(LineStyleType.Thin, Color.Black);

                workbook.SaveToFile(fullFileName + "report.xlsx", FileFormat.Version2016);
                MessageBox.Show(owner, "Report has been created successfully.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                PrintFile(owner, fullFileName + "report.xlsx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, string.Format("An error has occurred while creating document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void CreateFile(Window owner)
        {            
            try
            {
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                DataTable data = GetData();
                ExcelFont fontBold = workbook.CreateFont();

                sheet.PageSetup.Orientation = PageOrientationType.Landscape;
                fontBold.IsBold = true;

                var headers = data.AsEnumerable()
                            .Select(s => s.Field<string>("Name"))
                            .ToArray();

                var cellsData = data.AsEnumerable()
                            .Select(s => s.Field<int>("Type"))
                            .ToArray();

                int columnsRange = 2 + headers.Length - 1;

                sheet.InsertArray(headers, 2, 2, false);
                sheet.InsertArray(cellsData, 3, 2, false);

                CellRange range = sheet.Range[2, 2, 2, columnsRange];
                
                range.Style.Rotation = 60;
                range.Style.Font.IsBold = true;
                range.AutoFitRows();
                range.ColumnWidth = 4;

                sheet.Range[2, 2, 3, columnsRange].BorderInside(LineStyleType.Thin, Color.Black);
                sheet.Range[2, 2, 3, columnsRange].BorderAround(LineStyleType.Thin, Color.Black);

                workbook.SaveToFile(FileName, GetFileFormat());
                System.Windows.MessageBox.Show(owner, "File has been created successfully.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);

                workbook.SaveToFile("file.xps", FileFormat.XPS);
                PrintDialog dlg = new PrintDialog();
                XpsDocument xpsDoc = new XpsDocument(@"file.xps", System.IO.FileAccess.Read);
                dlg.PrintDocument(xpsDoc.GetFixedDocumentSequence().DocumentPaginator, "Document title");
                //PrintFile(owner, "file.xps");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(owner, string.Format("An error has occurred while creating document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public PrinterSettings.StringCollection GetAvailablePrinters()
        {
            return PrinterSettings.InstalledPrinters;
        }

        private DataTable GetData()
        {
            DataTable _dt = new DataTable();

            _dt.Columns.Add("Name", Type.GetType("System.String"));
            _dt.Columns.Add("Type", Type.GetType("System.Int32"));

            _dt.Rows.Add(new Object[] { "Demerol 25MG/ML", 9 });
            _dt.Rows.Add(new Object[] { "Demerol 50MG/ML", 10 });
            _dt.Rows.Add(new Object[] { "Duramorph 5MG/ML", 8 });
            _dt.Rows.Add(new Object[] { "Ephedrin Injection 50MG/ML", 2 });
            _dt.Rows.Add(new Object[] { "Fentanyl 100MCG/2ML", 1 });
            _dt.Rows.Add(new Object[] { "Fentanyl 250MCG/5ML", 3 });
            _dt.Rows.Add(new Object[] { "Ketamine HCL 25MG/ML", 4 });
            _dt.Rows.Add(new Object[] { "Pentothal 25MG/ML", 1 });
            _dt.Rows.Add(new Object[] { "Versed 5MG/ML", 6 });
            _dt.Rows.Add(new Object[] { "Versed 25MG/ML", 7 });

            return _dt;
        }

            public void PrintFile(Window owner, string fullFileName)
        {
            /*if (string.IsNullOrEmpty(PrinterName) && string.IsNullOrWhiteSpace(PrinterName))
            {
                System.Windows.MessageBox.Show(owner, "Please select your printer name.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {*/
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(string.IsNullOrEmpty(fullFileName) ? FileName : fullFileName);
                System.Windows.Forms.PrintDialog dialog = new System.Windows.Forms.PrintDialog();
                dialog.AllowPrintToFile = true;
                dialog.AllowCurrentPage = true;
                dialog.AllowSomePages = true;
                dialog.AllowSelection = true;
                dialog.UseEXDialog = true;
                dialog.PrinterSettings.Duplex = Duplex.Simplex;
                dialog.PrinterSettings.FromPage = 0;
                dialog.PrinterSettings.ToPage = 8;
                dialog.PrinterSettings.PrintRange = PrintRange.SomePages;
                workbook.PrintDialog = dialog;
                PrintDocument pd = workbook.PrintDocument;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    pd.Print();
            //}
        }

        private FileFormat GetFileFormat()
        {
            if (FileName.EndsWith(".xls"))
                return FileFormat.Version97to2003;
            else if (FileName.EndsWith(".xlsx"))
                return FileFormat.Version2016;
            else
                return FileFormat.Version2013;
        }
    }
}
