#region Usings
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.Data;
using System.Drawing.Printing;
using System.IO;
using System.Windows;
using static System.Drawing.Printing.PrinterSettings;
#endregion

namespace AsposeDemo
{
    public class AsposeExcel : IExcelProcessor
    {
        #region Variables
        public string FileName { get; set; }
        public string PrinterName { get; set; }
        #endregion

        public void PrintFile(Window owner, string fullFileName)
        {
            Workbook workbook = new Workbook(string.IsNullOrEmpty(fullFileName) ? FileName : fullFileName);

            if (string.IsNullOrEmpty(PrinterName) && string.IsNullOrWhiteSpace(PrinterName))
            {
                MessageBox.Show(owner, "Please select your printer name.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                // Apply different Image/Print options.
                Aspose.Cells.Rendering.ImageOrPrintOptions options = new Aspose.Cells.Rendering.ImageOrPrintOptions();
                options.ImageFormat = System.Drawing.Imaging.ImageFormat.Png;
                options.PrintingPage = PrintingPageType.Default;

                // To print a whole workbook, iterate through the sheets and print them, or use the WorkbookRender class.
                WorkbookRender wr = new WorkbookRender(workbook, options);

                // Print the workbook.
                try
                {
                    // Setting the number of pages to which the width of the worksheet will be spanned
                    wr.ToPrinter(PrinterName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(owner, string.Format("An error has occurred while printing document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        public void CreateFile(Window owner)
        {
            try
            {
                Workbook wb = new Workbook();
                wb.Worksheets[0].PageSetup.Orientation = PageOrientationType.Landscape;

                DataTable data = GetData();

                var column = 66; // Letter B in ASCII
                foreach (DataRow dr in data.Rows)
                {
                    char columnChar = (char)column;
                    Cell header = wb.Worksheets[0].Cells[string.Format("{0}{1}", columnChar, 2)];
                    Aspose.Cells.Style objstyle = header.GetStyle();

                    // Specify the angle of rotation of the text.
                    objstyle.RotationAngle = 60;
                    objstyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    objstyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    objstyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    objstyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    header.PutValue(dr[0]);
                    objstyle.Pattern = BackgroundType.Solid;

                    header.SetStyle(objstyle);

                    Cell cell = wb.Worksheets[0].Cells[string.Format("{0}{1}", columnChar, 3)];
                    Aspose.Cells.Style cellStyle = cell.GetStyle();

                    cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cell.PutValue(dr[1]);
                    cell.SetStyle(cellStyle);

                    column++;
                }

                for(var i = 1; i <= data.Rows.Count; i++)
                    wb.Worksheets[0].Cells.SetColumnWidth(i, 4);

                Cell cosignatureCell = wb.Worksheets[0].Cells[string.Format("{0}{1}", (char)((column + 1)), 2)];
                cosignatureCell.PutValue("COSIGNATURE FOR WASTE");
                Aspose.Cells.Style cosignatureStyle = cosignatureCell.GetStyle();
                cosignatureStyle.IsTextWrapped = true;
                cosignatureCell.SetStyle(cosignatureStyle);
                
                wb.Worksheets[0].AutoFitRow(1);
                //Save the Shared Workbook
                wb.Save(FileName);
                MessageBox.Show(owner, "File has been created successfully.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
               PrintFile(owner, FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, string.Format("An error has occurred while creating document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void BuildReport(Window owner)
        {
            if (string.IsNullOrEmpty(PrinterName) && string.IsNullOrWhiteSpace(PrinterName))
            {
                MessageBox.Show(owner, "Please select your printer name.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            FileStream stream = null;
            try
            {
                string dataDir = System.IO.Directory.GetCurrentDirectory();
                string fullFileName = dataDir + @"\templates\";

                stream = new FileStream(fullFileName + "Template.xlsx", FileMode.Open);

                // Instantiate LoadOptions specified by the LoadFormat.
                //LoadOptions loadOptions1 = new LoadOptions(LoadFormat.Excel97To2003);

                // Create a Workbook object and opening the file from the stream
                Workbook wb = new Workbook(stream);//, loadOptions1);
                wb.Worksheets[0].PageSetup.Orientation = PageOrientationType.Landscape;

                DataTable data = GetData();

                var column = 67; // Letter C in ASCII
                var columnNumber = 3;
                foreach (DataRow dr in data.Rows)
                {
                    char columnChar = (char)column;
                    Cell header = wb.Worksheets[0].Cells[string.Format("{0}{1}", columnChar, 17)];
                    Aspose.Cells.Style objstyle = header.GetStyle();

                    // Specify the angle of rotation of the text.
                    objstyle.RotationAngle = 60;
                    objstyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    objstyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    objstyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    objstyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    header.PutValue(dr[0]);
                    objstyle.Pattern = BackgroundType.Solid;
                    header.SetStyle(objstyle);

                    Cell cell = wb.Worksheets[0].Cells[string.Format("{0}{1}", columnChar, 18)];
                    Aspose.Cells.Style cellStyle = cell.GetStyle();

                    cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cell.PutValue(dr[1]);
                    cell.SetStyle(cellStyle);

                    wb.Worksheets[0].AutoFitColumn(columnNumber);

                    column++;
                    columnNumber++;
                }

                for (var i = 2; i <= data.Rows.Count; i++)
                    wb.Worksheets[0].Cells.SetColumnWidth(i, 4);

                wb.Worksheets[0].AutoFitRows(true);
                //Save the Shared Workbook
                wb.Save(fullFileName + "report.xlsx");
                MessageBox.Show(owner, "Report has been created successfully.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                PrintFile(owner, fullFileName + "report.xlsx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, string.Format("An error has occurred while creating document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (!object.Equals(null, stream))
                    stream.Close();
            }
        }

        public StringCollection  GetAvailablePrinters()
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
    }
}
