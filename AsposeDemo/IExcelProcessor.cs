using static System.Drawing.Printing.PrinterSettings;
using System.Windows;

namespace AsposeDemo
{
    public interface IExcelProcessor
    {
        string FileName { get; set; }
        string FileFormatExtension { get; set; }
        string PrinterName { get; set; }

        StringCollection GetAvailablePrinters();
        void PrintFile(Window owner, string fullFileName);
        void CreateFile(Window owner);
        void BuildReport(Window owner);
        void GetPdf(Window owner);
    }
}
