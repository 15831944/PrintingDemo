#region Usings
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using System.Diagnostics;
using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using ZXing;
using static System.Drawing.Printing.PrinterSettings;
using System.Collections.Generic;
#endregion

namespace AsposeDemo
{
    public class AsposeExcel : IExcelProcessor
    {
        #region Variables
        public string FileName { get; set; }
        public string PrinterName { get; set; }
        public string FileFormatExtension { get; set; }
        private bool Convert2Pdf { get; set; }

        public Dictionary<string, List<string>> TimesElapsed = new Dictionary<string, List<string>>(); 
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
            Stopwatch myTimer = new Stopwatch();

            if (string.IsNullOrEmpty(PrinterName) && string.IsNullOrWhiteSpace(PrinterName))
            {
                MessageBox.Show(owner, "Please select your printer name.", "Atention", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            FileStream stream = null;
            try
            {
                string txtTransactionno1 = "AN11160002";
                string filledBy = "KIT WAS USED FROM DATE 05 / 01 / 2018   TO DATE @tMonth / @tDay / @tYear";
                string dataDir = System.IO.Directory.GetCurrentDirectory();
                string fullFileName = dataDir + @"\templates\";
                var date = System.DateTime.Now.ToString("MM-dd-yyyy");
                var isOldExcel = FileName.EndsWith("xls");

                myTimer.Start();

                stream = new FileStream(fullFileName + FileName, FileMode.Open);

                // Instantiate LoadOptions specified by the LoadFormat.
                LoadOptions loadOptions1 = new LoadOptions(LoadFormat.Excel97To2003);

                // Create a Workbook object and opening the file from the stream
                Workbook wb = isOldExcel ? new Workbook(stream, loadOptions1) : new Workbook(stream);
                wb.Worksheets[0].PageSetup.Orientation = PageOrientationType.Landscape;

                TextBoxCollection textBoxes = wb.Worksheets[0].TextBoxes;
                textBoxes["txtTransactionno1"].Text = txtTransactionno1;
                textBoxes["txtTransactionno1"].Font.IsBold = true;

                textBoxes["txthospitalname"].Text = "General Hospital";
                textBoxes["txthospitalname"].Fill.FillType = FillType.Solid;
                textBoxes["txthospitalname"].Fill.SolidFill.Color = Color.White;

                textBoxes["txtsiteid"].Text = "San Diego, CA";
                textBoxes["txtsiteid"].Fill.FillType = FillType.Solid;
                textBoxes["txtsiteid"].Fill.SolidFill.Color = Color.White;

                var picture = GenerateBarCode(txtTransactionno1, 247, 38);

                wb.Worksheets[0].Pictures.Add(8, 0, picture);
                
                textBoxes["txtLocation"].Text = "LD1";
                textBoxes["txtLocation"].Font.IsBold = true;
                textBoxes["txttransactionno"].Text = txtTransactionno1;
                textBoxes["txttransactionno"].Font.IsBold = true;

                wb.Worksheets[0].Cells["B5"].Value = filledBy.Replace("@tMonth", date.Split('-')[0]).Replace("@tDay", date.Split('-')[1]).Replace("@tYear", date.Split('-')[2]);

                DataTable data = GetData();

                var column = 3;
                foreach (DataRow dr in data.Rows)
                {
                    //char columnChar = (char)column;
                    Cell header = wb.Worksheets[0].Cells[8, column];
                    Aspose.Cells.Style objstyle = header.GetStyle();

                    // Specify the angle of rotation of the text.
                    //objstyle.RotationAngle = 60;
                    objstyle.Font.Size = 10;
                    //objstyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                    header.PutValue(dr[0]);
                    header.SetStyle(objstyle);

                    Cell cell = wb.Worksheets[0].Cells[9, column - 1];

                    cell.PutValue(dr[1]);
                    column+=4;
                }

                //Save the Shared Workbook
                if (isOldExcel)
                    wb.Save(fullFileName + "report.xls", SaveFormat.Excel97To2003);
                else
                    wb.Save(fullFileName + "report.xlsx", SaveFormat.Xlsx);

                myTimer.Stop();
                AddTimeElapsed(FileFormatExtension, myTimer.Elapsed.ToString());

                if (Convert2Pdf)
                {
                    wb.Save(fullFileName + "report.pdf", SaveFormat.Pdf);
                    System.Diagnostics.Process.Start(fullFileName + "report.pdf");
                    Convert2Pdf = false;
                }
                else
                {
                    MessageBox.Show(owner, string.Format("Report has been created successfully."), "Information", MessageBoxButton.OK, MessageBoxImage.Information);

                    ExcelViewer ev = new ExcelViewer(ConvertSheetToImage(wb.Worksheets[0]));
                    ev.Closed += (object sender, EventArgs e) =>
                    {
                        PrintFile(owner, fullFileName + (isOldExcel ? "report.xls" : "report.xlsx"));
                    };
                    ev.Show();
                }
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

        public void GetPdf(Window owner)
        {
            try
            {
                Convert2Pdf = true;
                BuildReport(owner);
                MessageBox.Show(owner, string.Format("PDF Report has been created successfully."), "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, string.Format("An error has occurred while creating PDF document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Convert2Pdf = false;
            }
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

        private Stream GenerateBarCode(string value, int width, int height)
        {
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.CODE_39,
                Options = new ZXing.Common.EncodingOptions
                {
                    Height = height,
                    Width = width,
                    Margin = 1,
                    PureBarcode = true
                }
            };
            var bitmap = writer.Write(value);
            return GetStreamFromBitmap(bitmap);
        }

        private Stream GetStreamFromBitmap(Bitmap bitmapImage)
        {
            MemoryStream memStream = null;
            try
            {
                memStream = new MemoryStream();
                bitmapImage.Save(memStream, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            catch (Exception ex)
            {
                ;
            }
            return memStream;
        }

        private Bitmap ConvertSheetToImage(Worksheet sheet)
        {
            // Apply different Image/Print options.
            ImageOrPrintOptions options = new Aspose.Cells.Rendering.ImageOrPrintOptions();
            options.ImageFormat = System.Drawing.Imaging.ImageFormat.Png;
            options.PrintingPage = PrintingPageType.Default;

            SheetRender sr = new Aspose.Cells.Rendering.SheetRender(sheet, options);

            Bitmap bitmap = sr.ToImage(0);

            return bitmap;
        }

        private void AddTimeElapsed(string key, string value)
        {
            if (TimesElapsed.ContainsKey(key))
            {
                List<string> values = null;

                TimesElapsed.TryGetValue(key, out values);
                values.Add(value);
                TimesElapsed[key] = values;
            } else
            {
                List<string> values = new List<string>();
                values.Add(value);
                TimesElapsed.Add(key, values);
            }
        }
    }
}
