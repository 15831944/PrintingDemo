using GemBox.Spreadsheet;
using System;
using System.Linq;
using System.Data;
using System.Drawing.Printing;
using System.Windows;
using System.Drawing;
using System.Windows.Media;

namespace AsposeDemo.implementations
{
    public class GemboxSpreadSheetExcel : IExcelProcessor
    {
        private ExcelFile _ef { get; set; }
        public string FileName { get; set; }
        public string PrinterName { get; set; }

        public void BuildReport(Window owner)
        {
            try
            {
                // Set license key to use GemBox.Spreadsheet in Free mode.
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                string dataDir = System.IO.Directory.GetCurrentDirectory();
                string fullFileName = dataDir + @"\templates\";

                _ef = ExcelFile.Load(fullFileName + "Template.xlsx");
                var ws = _ef.Worksheets.ActiveWorksheet;

                DataTable data = GetData();

                var headers = data.AsEnumerable()
                                .Select(s => s.Field<string>("Name"))
                                .ToArray();

                var cellsData = data.AsEnumerable()
                            .Select(s => s.Field<int>("Type"))
                            .ToArray();

                CellStyle tmpStyle = new CellStyle();
                tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
                tmpStyle.Font.Weight = ExcelFont.BoldWeight;
                tmpStyle.Font.Color = System.Drawing.Color.Black;
                tmpStyle.Rotation = 60;

                int i = 2;
                // Write header data to Excel cells.
                foreach (var header in headers)
                {
                    ws.Cells[16, i].Value = header;
                    ws.Columns[i].AutoFit(1, ws.Rows[16], ws.Rows[17]);
                    ws.Cells[16, i].Style = tmpStyle;
                    ws.Cells[16, i].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                    i++;
                }

                tmpStyle = new CellStyle();
                tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
                tmpStyle.Font.Weight = ExcelFont.NormalWeight;

                i = 2;
                // Write data to Excel cells.
                foreach (var cell in cellsData)
                {
                    ws.Cells[17, i].Value = cell;
                    ws.Cells[17, i].Style = tmpStyle;
                    ws.Cells[17, i].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                    ws.Columns[i].Width = 4 * 256;
                    i++;
                }

                ws.Rows[1].AutoFit();
                ws.PrintOptions.FitWorksheetWidthToPages = 1;
                ws.PrintOptions.Portrait = false;

                _ef.Save(fullFileName + "report.xlsx");
                System.Windows.MessageBox.Show(owner, "File has been created successfully.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);

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
                // Set license key to use GemBox.Spreadsheet in Free mode.
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                _ef = new ExcelFile();
                ExcelWorksheet ws = _ef.Worksheets.Add("Report");

                DataTable data = GetData();

                var headers = data.AsEnumerable()
                                .Select(s => s.Field<string>("Name"))
                                .ToArray();

                var cellsData = data.AsEnumerable()
                            .Select(s => s.Field<int>("Type"))
                            .ToArray();

                CellStyle tmpStyle = new CellStyle();
                tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
                tmpStyle.Font.Weight = ExcelFont.BoldWeight;
                tmpStyle.Font.Color = System.Drawing.Color.Black;
                tmpStyle.Rotation = 60;

                int i = 1;
                // Write header data to Excel cells.
                foreach (var header in headers)
                {
                    ws.Cells[1, i].Value = header;
                    ws.Columns[i].AutoFit(1, ws.Rows[1], ws.Rows[2]);
                    ws.Cells[1, i].Style = tmpStyle;
                    ws.Cells[1, i].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                    i++;
                }

                tmpStyle = new CellStyle();
                tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
                tmpStyle.Font.Weight = ExcelFont.NormalWeight;

                i = 1;
                // Write data to Excel cells.
                foreach (var cell in cellsData)
                {
                    ws.Cells[2, i].Value = cell;
                    ws.Cells[2, i].Style = tmpStyle;
                    ws.Cells[2, i].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                    ws.Columns[i].Width = 4 * 256;
                    i++;
                }                

                ws.Rows[1].AutoFit();
                ws.PrintOptions.FitWorksheetWidthToPages = 1;
                ws.PrintOptions.Portrait = false;

                _ef.Save(FileName);
                System.Windows.MessageBox.Show(owner, "File has been created successfully.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);

                ExcelViewer ev = new ExcelViewer(_ef.ConvertToImageSource(ImageSaveOptions.ImageDefault));
                ev.Show();

                PrintFile(owner, string.Empty);
            } catch(Exception ex)
            {
                System.Windows.MessageBox.Show(owner, string.Format("An error has occurred while creating document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            } finally
            {
                if (!object.Equals(null, _ef))
                    _ef = null;
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
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            try
            {
                _ef = ExcelFile.Load(string.IsNullOrEmpty(FileName) ? fullFileName : FileName);
                _ef.Print(PrinterName);
            } catch(Exception ex)
            {
                System.Windows.MessageBox.Show(owner, string.Format("An error has occurred while creating document: {0}", ex.Message), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
